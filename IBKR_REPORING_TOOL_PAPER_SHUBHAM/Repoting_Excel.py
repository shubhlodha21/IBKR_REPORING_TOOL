"""
Daily IBKR MIS report for ONE account. Each run:
  1. Pulls live data from TWS / IB Gateway: today's fills, open orders,
     today's completed orders (for trigger/limit) and the account summary.
  2. Downloads the Flex Query statement (should cover the last 30 days) to fill
     in and confirm earlier days. A Flex failure is logged, not fatal.
  3. Merges both into a per-day fill store, de-duplicated by IBKR execution ID:
        data/<ACCOUNT>/fills/YYYY-MM-DD.json
        data/<ACCOUNT>/orders_ledger.json
        data/<ACCOUNT>/state.json
        data/runs.log
  4. Writes reports/MIS_paper_<date>_Shubham.xlsx with the sheets
     Index, Dashboard, Pending Order, Open Position, All Trades, Trade Summary,
     covering every trade from DATA_FROM up to the moment of the run.

Usage:
    run.ps1 (Windows) or run.sh (Linux/Mac) - both load .env first.

.env settings:
  IBKR_TOKEN, IBKR_QUERY_ID   Flex Web Service token and query (required)
  IBKR_ACCOUNT                account to report on, e.g. DU4293588. Optional when
                              the Gateway login has exactly one account.
  IBKR_DATA_FROM              optional report start date (YYYY-MM-DD). When unset,
                              the date of the first run for the account is used.
"""

import os
import re
import sys
import json
import glob
import time
import threading
import datetime as dt
import xml.etree.ElementTree as ET
from collections import defaultdict

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from ibapi.client import EClient
from ibapi.wrapper import EWrapper
from ibapi.execution import ExecutionFilter

# ----------------------------------------------------------------------
# CONFIG  (set IBKR_TOKEN and IBKR_QUERY_ID as environment variables)
# ----------------------------------------------------------------------
TOKEN    = os.environ.get("IBKR_TOKEN", "")
QUERY_ID = os.environ.get("IBKR_QUERY_ID", "")

if not TOKEN or not QUERY_ID:
    sys.exit("Error: IBKR_TOKEN and IBKR_QUERY_ID must be set as environment variables.\n"
             "Copy .env.example to .env and fill in your credentials, then run via run.ps1.")

# Account to report on. Resolved at run time (see resolve_account) when unset.
ACCOUNT = os.environ.get("IBKR_ACCOUNT", "").strip()
_DATA_FROM_ENV = os.environ.get("IBKR_DATA_FROM", "").strip()

_BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
_REPORTS_DIR = os.path.join(_BASE_DIR, "reports")
OUTPUT_FILE  = os.path.join(_REPORTS_DIR, dt.date.today().strftime("MIS_paper_%d%b%Y") + "_Shubham.xlsx")

# Per-account store. Every fill is kept in the file for its trade date, keyed by
# IBKR execution ID, so Flex and live TWS copies of one fill share one record.
# Nothing is ever pruned, so a missed run or an empty Flex download can't shrink
# the report. ACCOUNT_DIR is set once the account is known (see set_account).
DATA_DIR    = os.path.join(_BASE_DIR, "data")
RUN_LOG     = os.path.join(DATA_DIR, "runs.log")
ACCOUNT_DIR = None

# Pre-2026-11 single-file stores. Imported once into the per-account store (only
# the chosen account's records) and then left untouched as a backup.
LEGACY_HISTORY_FILE = os.path.join(_BASE_DIR, "trade_history.json")
LEGACY_CACHE_FILE   = os.path.join(_BASE_DIR, "trades_cache.json")
LEGACY_LEDGER_FILE  = os.path.join(_BASE_DIR, "orders_ledger.json")

BASE     = "https://ndcdyn.interactivebrokers.com/AccountManagement/FlexWebService"
SEND_URL = f"{BASE}/SendRequest"
GET_URL  = f"{BASE}/GetStatement"
VERSION  = "3"

MAX_POLL_ATTEMPTS = 12
POLL_WAIT_SECONDS = 5

# TWS live connection
HOST       = "127.0.0.1"
PORT       = 4002       # 7496 live TWS | 7497 paper | 4001/4002 Gateway
CLIENT_ID  = 998
ORDER_WAIT = 8           # seconds to wait for open-order callbacks
ACCT_WAIT  = 8           # seconds to wait for account-summary callbacks
EXEC_WAIT  = 15          # seconds to wait for executions (0.5 s timer + fills + commissions)

_UNSET = (0, 1.7976931348623157e308)


def _sym_key(contract):
    """Symbol as the Flex statement writes it: FX pairs by their pair name
    (e.g. EUR.USD), everything else by its symbol."""
    if getattr(contract, "secType", "") == "CASH" and getattr(contract, "localSymbol", ""):
        return contract.localSymbol
    return contract.symbol


def _qty(v):
    """Quantity from the TWS API (float / Decimal / unset) as a plain number."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if f >= 1e300:                      # UNSET_DECIMAL / UNSET_DOUBLE
        return None
    return int(f) if f == int(f) else f

# Report start date — the report only includes trades on/after this date. Set at
# run time from IBKR_DATA_FROM, else the account's first run date (state.json).
# Shown on the Index and Dashboard so the reporting window is always explicit.
DATA_FROM = dt.date.today()

# Running positions use the same cutoff.
RUNNING_POSITIONS_FROM = DATA_FROM

# Risk limits shown on the Dashboard (hardcoded).
MAX_EXPOSURE   = 'NA'
DAILY_MAX_LOSS = 'NA'

# Cumulative loss attributed to bugs — a manually-maintained figure shown on the
# Dashboard. Update this value here when it changes; it stays fixed otherwise.
LOSS_DUE_TO_BUGS = 0

# IST offset for report timestamps
_IST = dt.timezone(dt.timedelta(hours=5, minutes=30))

# Account summary tags to request from TWS
ACCOUNT_SUMMARY_TAGS = (
    "NetLiquidation,TotalCashValue,SettledCash,AccruedCash,"
    "BuyingPower,AvailableFunds,ExcessLiquidity,"
    "InitMarginReq,MaintMarginReq,"
    "GrossPositionValue,UnrealizedPnL,RealizedPnL"
)

# Display labels for each account tag (label, tag_key)
ACCOUNT_DISPLAY = [
    ("Net Liquidation (Total Funds)", "NetLiquidation"),
    ("Total Cash Value",              "TotalCashValue"),
    ("Settled Cash",                  "SettledCash"),
    ("Accrued Cash",                  "AccruedCash"),
    ("Buying Power",                  "BuyingPower"),
    ("Available Funds",               "AvailableFunds"),
    ("Excess Liquidity",              "ExcessLiquidity"),
    ("Initial Margin Req.",           "InitMarginReq"),
    ("Maintenance Margin Req.",       "MaintMarginReq"),
    ("Gross Position Value",          "GrossPositionValue"),
    ("Unrealized PnL",                "UnrealizedPnL"),
    ("Realized PnL",                  "RealizedPnL"),
]


# ----------------------------------------------------------------------
# TWS: fetch live pending orders + account summary in one connection
# ----------------------------------------------------------------------
class IBOrderApp(EWrapper, EClient):
    def __init__(self):
        EClient.__init__(self, self)
        self.orders_by_symbol = defaultdict(list)
        self.account_data     = {}          # tag -> (value, currency)
        self.executions       = {}          # execId -> execution dict
        self.completed_orders = []          # today's filled/cancelled orders
        self.managed_accounts = []          # accounts on this Gateway login
        self.order_status     = {}          # permId -> (status, filled qty, avg fill price)
        self.portfolio        = {}          # symbol -> live position / price / P&L
        self.portfolio_account = None
        self._portfolio_done  = threading.Event()
        self._orders_done     = threading.Event()
        self._acct_done       = threading.Event()
        self._exec_done       = threading.Event()
        self._completed_done  = threading.Event()
        self.connected_ok     = False

    def error(self, reqId, errorCode, errorString, advancedOrderRejectJson=""):
        if errorCode in (502, 504, 1100, 1300):
            print(f"[TWS] {errorCode}: {errorString}", file=sys.stderr)
        elif errorCode not in (2104, 2106, 2107, 2119, 2158, 2100):
            print(f"[TWS] id={reqId} code={errorCode}: {errorString}", file=sys.stderr)

    def connectAck(self):
        self.connected_ok = True

    def managedAccounts(self, accountsList):
        self.managed_accounts = [a.strip() for a in accountsList.split(",") if a.strip()]

    def nextValidId(self, orderId):
        self.connected_ok = True
        self.reqAllOpenOrders()
        self.reqAccountSummary(9001, "All", ACCOUNT_SUMMARY_TAGS)
        # Delay executions request by 0.5 s so the first two requests
        # don't congest the message loop, and set a 24-hour window so
        # pre-market fills and all intraday trades are captured.
        threading.Timer(0.5, self._req_executions).start()
        # Open positions with IB's current market price and unrealized P&L.
        acct = ACCOUNT or (self.managed_accounts[0] if len(self.managed_accounts) == 1 else None)
        if acct:
            self.portfolio_account = acct
            self.reqAccountUpdates(True, acct)
        else:
            self._portfolio_done.set()

    def _req_executions(self):
        # Empty filter → ALL of the current trading day's executions.
        # NOTE: IBKR's API only ever returns the *current day's* executions;
        # a time filter cannot reach prior days, and a malformed one can
        # suppress results entirely — so we pass an unfiltered request.
        self.reqExecutions(10001, ExecutionFilter())
        # Today's completed (filled/cancelled) orders — these carry the order's
        # limit/trigger even after they've filled, which open orders no longer do.
        self.reqCompletedOrders(False)   # False = all, not just API-placed

    # ── completed orders (today's filled/cancelled, with limit/trigger) ──
    def completedOrder(self, contract, order, orderState):
        trigger = order.auxPrice if order.auxPrice not in _UNSET else None
        limit   = order.lmtPrice if order.lmtPrice not in _UNSET else None
        oid     = getattr(order, "permId", None) or getattr(order, "orderId", None)
        self.completed_orders.append({
            "orderId":   f"C{oid}",
            "permId":    getattr(order, "permId", 0) or None,
            "account":   getattr(order, "account", "") or "",
            "symbol":    _sym_key(contract),
            "action":    order.action,
            "orderType": order.orderType,
            "trigger":   trigger,
            "limit":     limit,
            "quantity":  int(order.totalQuantity) if order.totalQuantity else 0,
            "parentPermId": getattr(order, "parentPermId", 0) or None,
            "status":    orderState.status,
            "filled":    _qty(getattr(order, "filledQuantity", None)),
            "time":      getattr(orderState, "completedTime", "") or "",
        })

    def completedOrdersEnd(self):
        print(f"[TWS] completedOrdersEnd - {len(self.completed_orders)} completed order(s).")
        self._completed_done.set()

    # ── pending orders ────────────────────────────────────────────────
    def openOrder(self, orderId, contract, order, orderState):
        trigger = order.auxPrice if order.auxPrice not in _UNSET else None
        limit   = order.lmtPrice if order.lmtPrice not in _UNSET else None
        perm    = getattr(order, "permId", 0) or None
        legs    = self.orders_by_symbol[_sym_key(contract)]
        # TWS can send openOrder several times for one order (status updates);
        # keep only the latest copy.
        if perm:
            legs[:] = [o for o in legs if o.get("permId") != perm]
        legs.append({
            "orderId":   orderId,
            "permId":    getattr(order, "permId", 0) or None,
            "account":   getattr(order, "account", "") or "",
            "action":    order.action,
            "orderType": order.orderType,
            "trigger":   trigger,
            "limit":     limit,
            "quantity":  int(order.totalQuantity),
            "parentId":  order.parentId,
            "parentPermId": getattr(order, "parentPermId", 0) or None,
            "status":    getattr(orderState, "status", "") or "",
        })

    def orderStatus(self, orderId, status, filled, remaining, avgFillPrice, permId,
                    parentId, lastFillPrice, clientId, whyHeld, mktCapPrice=None):
        if permId:
            self.order_status[permId] = (status, _qty(filled), avgFillPrice)

    def openOrderEnd(self):
        self._orders_done.set()

    # ── account summary ───────────────────────────────────────────────
    def accountSummary(self, reqId, account, tag, value, currency):
        # Kept per account; main() picks the reported account's figures.
        self.account_data.setdefault(account, {})[tag] = (value, currency)

    def accountSummaryEnd(self, reqId):
        self._acct_done.set()

    # ── today's executions ────────────────────────────────────────────
    def execDetails(self, reqId, contract, execution):
        self.executions[execution.execId] = {
            "execId":     execution.execId,   # IBKR execution ID (== Flex ibExecID)
            "account":    execution.acctNumber,
            "permId":     getattr(execution, "permId", 0) or None,   # links fill -> order
            "conId":      getattr(contract, "conId", 0) or None,
            "contract":   _sym_key(contract),
            "action":     "BUY" if execution.side in ("BOT", "BUY") else "SELL",
            "quantity":   execution.shares,
            "price":      execution.price,
            "currency":   contract.currency,
            "exchange":   execution.exchange,
            "time":       execution.time,
            "commission": None,             # filled in by commissionReport
        }

    def commissionReport(self, commissionReport):
        ex = self.executions.get(commissionReport.execId)
        if ex is not None:
            ex["commission"] = commissionReport.commission

    def execDetailsEnd(self, reqId):
        print(f"[TWS] execDetailsEnd - {len(self.executions)} execution(s) received so far.")
        self._exec_done.set()

    # ── portfolio (open positions) ────────────────────────────────────
    def updatePortfolio(self, contract, position, marketPrice, marketValue,
                        averageCost, unrealizedPNL, realizedPNL, accountName):
        pos = _qty(position) or 0
        if not pos:
            return
        self.portfolio[_sym_key(contract)] = {
            "account":     accountName,
            "position":    pos,
            "marketPrice": marketPrice,
            "marketValue": marketValue,
            "averageCost": averageCost,
            "unrealized":  unrealizedPNL,
        }

    def accountDownloadEnd(self, accountName):
        print(f"[TWS] accountDownloadEnd - {len(self.portfolio)} open position(s).")
        self._portfolio_done.set()


def fetch_tws_data():
    """Connect to TWS once and fetch open orders, account summary, today's
    executions, and today's completed orders."""
    app = IBOrderApp()
    try:
        app.connect(HOST, PORT, CLIENT_ID)
    except Exception as e:
        print(f"[TWS] Could not connect ({e}) - live sheets will be empty.",
              file=sys.stderr)
        return {}, {}, {}, [], [], None
    threading.Thread(target=app.run, daemon=True).start()
    # Each feed has its own generous timeout so a slow/missing feed
    # doesn't block the others.
    app._orders_done.wait(timeout=ORDER_WAIT)
    app._acct_done.wait(timeout=ACCT_WAIT)
    app._exec_done.wait(timeout=EXEC_WAIT)
    app._completed_done.wait(timeout=EXEC_WAIT)
    got_portfolio = app._portfolio_done.wait(timeout=ACCT_WAIT)
    # Commission reports arrive asynchronously after execDetailsEnd;
    # give TWS extra time to push them all through before we disconnect.
    time.sleep(1.5)
    if app.portfolio_account:
        try:
            app.reqAccountUpdates(False, app.portfolio_account)
        except Exception:                                   # noqa: BLE001
            pass
    app.disconnect()
    # Attach the latest status / filled quantity to each open order.
    for legs in app.orders_by_symbol.values():
        for o in legs:
            st = app.order_status.get(o.get("permId"))
            if st:
                o["status"] = st[0] or o.get("status", "")
                o["filled"] = st[1]
    # None = portfolio not received (Gateway down) -> Unrealized P&L left blank.
    portfolio = app.portfolio if (got_portfolio and app.portfolio_account) else None
    return (app.orders_by_symbol, app.account_data, app.executions,
            app.completed_orders, app.managed_accounts, portfolio)


# ----------------------------------------------------------------------
# Small file helpers
# ----------------------------------------------------------------------
def _read_json(path, default):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            print(f"[store] {path} unreadable ({e}) - ignoring it.", file=sys.stderr)
    return default


def _write_json_atomic(path, obj):
    """Write via a temp file + rename, so a crash mid-write never leaves a
    half-written (corrupt) file behind."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, indent=2, sort_keys=True)
    os.replace(tmp, path)


def log_run(status, **info):
    """Append one line per run to data/runs.log."""
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        parts = " | ".join(f"{k}={v}" for k, v in info.items())
        with open(RUN_LOG, "a", encoding="utf-8") as f:
            f.write(f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} | {status} | "
                    f"account={ACCOUNT or '?'} | {parts}\n")
    except OSError as e:
        print(f"[run-log] could not write: {e}", file=sys.stderr)


def set_account(account):
    """Point every per-account path at data/<account>/."""
    global ACCOUNT, ACCOUNT_DIR
    ACCOUNT     = account
    ACCOUNT_DIR = os.path.join(DATA_DIR, account)


def _fills_dir():
    return os.path.join(ACCOUNT_DIR, "fills")


def _ledger_file():
    return os.path.join(ACCOUNT_DIR, "orders_ledger.json")


def _state_file():
    return os.path.join(ACCOUNT_DIR, "state.json")


def load_state():
    return _read_json(_state_file(), {})


def save_state(state):
    _write_json_atomic(_state_file(), state)


def set_data_from(d):
    global DATA_FROM, RUNNING_POSITIONS_FROM
    DATA_FROM = RUNNING_POSITIONS_FROM = d


# ----------------------------------------------------------------------
# Persistent order ledger: capture trigger/limit from live open orders so
# trades can later be matched to the prices set when the order was placed.
# (IBKR does NOT retain these on executed trades, so we capture them ourselves.)
# ----------------------------------------------------------------------
def _norm_type(t):
    """Normalise order type for matching (TWS 'STP LMT' vs Flex 'STPLMT')."""
    return (t or "").upper().replace(" ", "")


def load_order_ledger():
    return _read_json(_ledger_file(), {})


def update_order_ledger(ledger, orders_by_symbol, completed_orders=None):
    """Merge live open orders AND today's completed orders into the ledger.

    Open orders cover still-working orders; completed orders cover ones that
    already filled today (which an EOD run would otherwise miss). Both carry
    the limit/trigger prices set when the order was placed. Orders are keyed by
    permId (IBKR's permanent order ID, also carried on every fill) when known.
    """
    today = dt.date.today().isoformat()

    def _merge(o, symbol):
        perm = o.get("permId")
        oid  = f"P{perm}" if perm else str(o["orderId"])
        rec = ledger.get(oid, {})
        rec.update({
            "symbol":    symbol,
            "action":    (o.get("action") or "").upper(),
            "orderType": _norm_type(o.get("orderType")),
            "limit":     o.get("limit"),
            "trigger":   o.get("trigger"),
            "quantity":  o.get("quantity"),
            "last_seen": today,
        })
        if perm:
            rec["permId"] = perm
        if o.get("parentPermId"):
            rec["parentPermId"] = o["parentPermId"]
        rec.setdefault("first_seen", today)
        ledger[oid] = rec

    link_parent_orders(orders_by_symbol)
    for symbol, legs in orders_by_symbol.items():
        for o in legs:
            _merge(o, symbol)
    for o in (completed_orders or []):
        _merge(o, o["symbol"])
    return ledger


def link_parent_orders(orders_by_symbol):
    """Child orders sometimes only carry parentId (the parent's client order
    ID); translate it to the parent's permId using the open orders seen."""
    perm_by_oid = {o["orderId"]: o.get("permId")
                   for legs in orders_by_symbol.values() for o in legs
                   if o.get("orderId") and o.get("permId")}
    for legs in orders_by_symbol.values():
        for o in legs:
            if not o.get("parentPermId") and o.get("parentId"):
                o["parentPermId"] = perm_by_oid.get(o["parentId"])


def save_order_ledger(ledger):
    _write_json_atomic(_ledger_file(), ledger)


class OrderLookup:
    """Ledger index: exact match by permId, else best guess by
    (symbol, action, order type) and the date the order was last seen."""

    def __init__(self, ledger):
        self.by_perm   = {}
        self.children  = defaultdict(list)   # parent permId -> child order records
        self.by_key    = defaultdict(list)
        for rec in ledger.values():
            if rec.get("permId"):
                self.by_perm[str(rec["permId"])] = rec
            if rec.get("parentPermId"):
                self.children[str(rec["parentPermId"])].append(rec)
            key = (rec.get("symbol"), (rec.get("action") or "").upper(),
                   _norm_type(rec.get("orderType")))
            try:
                seen = dt.date.fromisoformat(rec.get("last_seen", ""))
            except (ValueError, TypeError):
                seen = dt.date.min
            self.by_key[key].append((seen, rec.get("limit"), rec.get("trigger")))


def build_order_lookup(ledger):
    return OrderLookup(ledger)


def lookup_order_prices(order_lookup, symbol, action, order_type, trade_date, perm_id=None):
    """Return (limit, trigger) for the ledger order behind a trade."""
    if not order_lookup:
        return None, None
    if perm_id:
        rec = order_lookup.by_perm.get(str(perm_id))
        if rec:
            return rec.get("limit"), rec.get("trigger")
    key = (symbol, (action or "").upper(), _norm_type(order_type))
    cands = order_lookup.by_key.get(key, [])
    if not cands:
        return None, None
    td = trade_date or dt.date.min
    prior = [c for c in cands if c[0] <= td]           # orders seen on/before the trade
    pick = (max(prior, key=lambda c: c[0]) if prior
            else min(cands, key=lambda c: abs((c[0] - td).days)))
    return pick[1], pick[2]


# ----------------------------------------------------------------------
# Fill store: data/<ACCOUNT>/fills/YYYY-MM-DD.json
#
# One record per fill, keyed by IBKR execution ID. Flex (`ibExecID`) and the
# live TWS feed (`execId`) report the SAME ID for the same fill, so a fill
# captured live today and published by Flex tomorrow lands on ONE record,
# tagged source=BOTH. Different IDs are genuinely different fills and are all
# kept, even when date/side/qty/price are identical.
#
# Corrections: IBKR re-issues a corrected fill with the last part of the ID
# bumped (e.g. ...01.01 -> ...01.02). The key is the ID without that last part,
# and the highest version wins.
# ----------------------------------------------------------------------
def _exec_id(row):
    for k in ("ibExecID", "execID", "execId"):
        v = row.get(k)
        if v not in (None, ""):
            return str(v).strip()
    return None


def _exec_base_and_version(eid):
    parts = eid.split(".")
    if len(parts) >= 3:
        try:
            return ".".join(parts[:-1]), int(parts[-1], 16)
        except ValueError:
            pass
    return eid, 0


def _fill_dedup_key(symbol, action, qty, price, d):
    """Fallback identity for the few legacy fills that carry no execution ID.
    Quantity is unsigned: Flex reports sells as negative, TWS as positive."""
    a = str(action).upper().strip()
    a = "BUY" if a in ("BUY", "BOT", "B") else "SELL" if a in ("SELL", "SLD", "S") else a
    return (d, str(symbol), a, round(abs(_flt(qty)), 4), round(_flt(price), 6))


def _fill_key(row):
    eid = _exec_id(row)
    if eid:
        return "id:" + _exec_base_and_version(eid)[0]
    k = _fill_dedup_key(row.get("symbol") or row.get("description"), row.get("buySell"),
                        row.get("quantity"), row.get("tradePrice"), parse_trade_date(row))
    return "fp:" + "|".join(str(p) for p in k)


def _fill_day(row):
    d = parse_trade_date(row)
    return d.isoformat() if d else "undated"


def load_fill_store():
    """All stored fills for the account: {key: record}."""
    store = {}
    for path in sorted(glob.glob(os.path.join(_fills_dir(), "*.json"))):
        store.update(_read_json(path, {}))
    return store


def save_fill_store(store):
    """Write one file per trade date. Only files whose content changed are
    rewritten; a day file left empty (its fill moved date) is removed."""
    by_day = defaultdict(dict)
    for key, rec in store.items():
        by_day[_fill_day(rec)][key] = rec
    existing = {os.path.basename(p)[:-5]: p
                for p in glob.glob(os.path.join(_fills_dir(), "*.json"))}
    written = 0
    for day, recs in by_day.items():
        path = os.path.join(_fills_dir(), f"{day}.json")
        if _read_json(path, None) != recs:
            _write_json_atomic(path, recs)
            written += 1
    for day, path in existing.items():
        if day not in by_day:
            os.remove(path)
    return written


def upsert_fill(store, row, source):
    """Insert or merge one fill. Returns 'new', 'merged' or 'skipped'."""
    clean = {k: v for k, v in row.items() if not k.startswith("_")}
    key   = _fill_key(clean)
    eid   = _exec_id(clean)

    # A legacy no-ID record for this same fill is replaced by the ID record.
    if eid:
        k = _fill_dedup_key(clean.get("symbol") or clean.get("description"),
                            clean.get("buySell"), clean.get("quantity"),
                            clean.get("tradePrice"), parse_trade_date(clean))
        legacy = store.pop("fp:" + "|".join(str(p) for p in k), None)
        if legacy and key not in store:
            store[key] = legacy

    old = store.get(key)
    if old is None:
        clean["source"]     = source
        clean["first_seen"] = dt.date.today().isoformat()
        store[key] = clean
        return "new"

    merged = dict(old)
    old_eid = _exec_id(old)
    new_ver = _exec_base_and_version(eid)[1] if eid else 0
    old_ver = _exec_base_and_version(old_eid)[1] if old_eid else 0
    if eid and old_eid and new_ver < old_ver:
        return "skipped"                    # older version of a corrected fill
    # Flex is the official record: once a fill has Flex data, a later live copy
    # only fills blanks (e.g. it must not replace the full name with the ticker).
    flex_owned = (source == "LIVE" and old.get("source") in ("FLEX", "BOTH")
                  and not new_ver > old_ver)
    for k, v in clean.items():
        if v in (None, ""):                 # never blank out a known value
            continue
        if flex_owned and merged.get(k) not in (None, ""):
            continue
        merged[k] = v
    srcs = {old.get("source") or source, source}
    merged["source"] = "BOTH" if ("BOTH" in srcs or {"LIVE", "FLEX"} <= srcs) else source
    if eid and old_eid and new_ver > old_ver:
        merged["corrected"] = True
    store[key] = merged
    return "merged"


def import_legacy_stores(store, ledger):
    """One-time import of the old single-file stores (trade_history.json,
    trades_cache.json, orders_ledger.json), keeping ONLY the reported account's
    fills. The old files are left in place as a backup."""
    n = 0
    for rec in _read_json(LEGACY_HISTORY_FILE, {}).values():
        if rec.get("accountId") != ACCOUNT:
            continue
        src = "FLEX" if rec.get("conid") not in (None, "") else "LIVE"
        upsert_fill(store, rec, src)
        n += 1
    for exec_id, ex in _read_json(LEGACY_CACHE_FILE, {}).items():
        if ex.get("account") != ACCOUNT:
            continue
        row = _exec_to_flex_row(ex)
        row["ibExecID"] = row.get("ibExecID") or exec_id
        upsert_fill(store, row, "LIVE")
        n += 1
    m = 0
    for oid, rec in _read_json(LEGACY_LEDGER_FILE, {}).items():
        ledger.setdefault(oid, rec)
        m += 1
    print(f"  [Import] {n} fill record(s) and {m} order(s) imported from the old "
          f"single-file stores for {ACCOUNT}.")


# ----------------------------------------------------------------------
# Pending order helpers
# ----------------------------------------------------------------------
def _fmt_price(v):
    return f"{v:.2f}" if v is not None else ""


# SL % shown on every BUY row of Pending Order and All Trades Detail.
BUY_SL_PCT = "0.25%"


def _sl_cell(action, computed=""):
    """SL % cell: the fixed BUY_SL_PCT on BUY rows, else the computed value."""
    return BUY_SL_PCT if str(action).upper() == "BUY" else computed


def _sl_pct(entry_trigger, sl_trigger):
    if entry_trigger and sl_trigger:
        return f"{abs(entry_trigger - sl_trigger) / entry_trigger * 100:.2f}%"
    return ""


def _offset(trigger, limit_price):
    if trigger is not None and limit_price is not None:
        return f"{abs(limit_price - trigger):.2f}"
    return ""


def _total_amount(qty, trigger, limit_price):
    price = trigger if trigger is not None else limit_price
    if price is not None:
        return f"{qty * price:,.2f}"
    return ""


def build_pending_rows(orders_by_symbol):
    """One numbered block per working entry order (entry row + its stop/exit
    rows). When the entry has already filled - the position is open and only
    its stop is still working - the stop is shown once, as its own block.
    Name is blank: TWS orders don't carry the Flex full name."""
    rows = []
    sr = 0

    def _is_child_of(child, parent):
        return ((child.get("parentPermId") and child.get("parentPermId") == parent.get("permId"))
                or (child.get("parentId") and child.get("parentId") == parent.get("orderId")))

    for symbol, legs in orders_by_symbol.items():
        parents  = [o for o in legs if not (o.get("parentId") or o.get("parentPermId"))]
        children = [o for o in legs if o not in parents]
        used = set()

        for entry in parents:
            kids = [c for c in children if _is_child_of(c, entry)]
            used.update(id(c) for c in kids)
            stop = next((c for c in kids if _norm_type(c.get("orderType")).startswith("STP")),
                        kids[0] if kids else None)
            sr += 1
            # Entry row - with SL % and Total Amount.
            rows.append([
                sr, symbol, "", entry["action"], entry["orderType"],
                _fmt_price(entry["trigger"]), _fmt_price(entry["limit"]),
                _offset(entry["trigger"], entry["limit"]),
                _sl_cell(entry["action"], _sl_pct(entry["trigger"] or entry["limit"],
                                                  stop["trigger"] if stop else None)),
                entry["quantity"],
                _total_amount(entry["quantity"], entry["trigger"], entry["limit"]),
            ])
            # Its stop / exit rows - no Sr No, SL % or Total Amount.
            for c in kids:
                rows.append([
                    "", symbol, "", c["action"], c["orderType"],
                    _fmt_price(c["trigger"]), _fmt_price(c["limit"]),
                    _offset(c["trigger"], c["limit"]),
                    _sl_cell(c["action"]), c["quantity"], "",
                ])

        # Children whose entry already filled (open position): one row each.
        for c in children:
            if id(c) in used:
                continue
            sr += 1
            rows.append([
                sr, symbol, "", c["action"], c["orderType"],
                _fmt_price(c["trigger"]), _fmt_price(c["limit"]),
                _offset(c["trigger"], c["limit"]),
                _sl_cell(c["action"]), c["quantity"],
                _total_amount(c["quantity"], c["trigger"], c["limit"]),
            ])
    return rows


PENDING_HEADERS = ["Sr No", "Contract", "Name", "Action", "Type",
                   "Trigger", "Limit", "Offset", "SL-Percentage", "Quantity", "Total Amount",
                   "Realized P&L", "Unrealized P&L"]
PENDING_WIDTHS  = [7, 12, 30, 9, 10, 10, 10, 10, 15, 10, 14, 14, 15]


def add_pnl_to_pending(pending_rows, agg_rows, portfolio):
    """Append the contract's Realized P&L (report window) and Unrealized P&L
    (live open position) to each numbered Pending Order row."""
    realized = {r["Contract"]: r["Realized P&L"] for r in agg_rows}
    out = []
    for row in pending_rows:
        if row[0] == "":
            out.append(row + ["", ""])
            continue
        sym = row[1]
        unreal = ("" if portfolio is None else
                  round(_flt((portfolio.get(sym) or {}).get("unrealized")), 2))
        out.append(row + [realized.get(sym, 0.0), unreal])
    return out

TODAY_HEADERS = ["Date & Time (UTC)", "Date & Time (GST)", "Sr No", "Contract", "Name",
                 "Action", "Quantity", "Price", "Trigger Price", "Limit Price",
                 "Exchange", "Order Type", "Commission", "Realized P&L", "Unrealized P&L",
                 "Account"]
TODAY_WIDTHS  = [22, 22, 7, 14, 30, 10, 12, 14, 14, 14, 12, 12, 14, 14, 15, 14]

# Gulf Standard Time (Dubai) is UTC+4 — used for the second All Trades time column.
_GST_OFFSET_HOURS = 4

# Columns whose numeric cells get green (positive) / red (negative) shading.
# Commission is intentionally excluded — it is shown with no colour.
COLOR_COLUMNS_TRADES = ["Realized P&L", "Unrealized P&L"]


def _parse_exec_time(raw):
    """Parse IBKR execution.time string into a display-friendly datetime."""
    raw        = str(raw).strip()
    normalised = " ".join(raw.split())
    for fmt in ("%Y%m%d %H:%M:%S", "%Y%m%d"):
        try:
            return dt.datetime.strptime(normalised, fmt).strftime("%d-%b-%Y %H:%M:%S")
        except ValueError:
            continue
    return raw


def _to_gst(utc_str):
    """Convert a formatted UTC datetime string to Gulf Standard Time (Dubai, UTC+4).
    Returns the shifted 'DD-Mon-YYYY HH:MM:SS' string, or the original unchanged
    when there is no parseable time component to shift."""
    if not utc_str:
        return ""
    try:
        d = dt.datetime.strptime(str(utc_str), "%d-%b-%Y %H:%M:%S")
    except ValueError:
        return utc_str   # date-only or unparseable — leave as-is
    return (d + dt.timedelta(hours=_GST_OFFSET_HOURS)).strftime("%d-%b-%Y %H:%M:%S")


def _trade_price_key(dt_str, symbol, action, qty_str, price_str):
    """Stable identity for a single fill, used to carry forward manual prices.
    Built from the exact formatted strings written to the All Trades sheet."""
    return (str(dt_str), str(symbol), str(action), str(qty_str), str(price_str))


def build_individual_trade_rows(trade_rows, date_filter=None, order_lookup=None,
                                manual_prices=None):
    """
    Build per-trade rows from Flex Query data.

    date_filter   : a collection of dates to include (e.g. {today, yesterday}).
                    If None, ALL trades are included (full history).
    order_lookup  : optional ledger lookup (see build_order_lookup) used to fill
                    the Trigger / Limit columns from previously-captured orders.
    manual_prices : optional {trade_key: (trigger, limit)} carried forward from
                    the previous report, where the user typed prices by hand.
                    These take precedence over the ledger / Flex values.
    """
    if date_filter is not None:
        sel = [r for r in trade_rows if parse_trade_date(r) in date_filter]
    else:
        sel = list(trade_rows)

    # Newest first, by the real trade time.
    sel.sort(key=_trade_sort_dt, reverse=True)

    rows = []
    for sr, r in enumerate(sel, start=1):
        f = _fill_fields(r, order_lookup, manual_prices)
        rows.append([
            f["dt"],            # Date & Time (UTC)
            f["dt_gst"],        # Date & Time (GST) — Dubai, UTC+4
            sr,
            f["symbol"],
            f["name"],
            f["action"],
            f["qty"],
            f["price"],
            f["trigger"],
            f["limit"],
            f["exchange"],
            r.get("orderType", ""),
            f["commission"],
            f["realized"],
            f["unrealized"],
            f["account"],
        ])
    return rows


def _fmt_abs(v, dp=2):
    """Absolute-value format (for qty/price/commission)."""
    try:
        return f"{abs(float(v)):.{dp}f}" if v not in (None, "") else ""
    except (ValueError, TypeError):
        return str(v) if v else ""


def _amt(v, signed=False):
    """Numeric amount (so the cell can carry the 1,234.00 number format).
    Returns a rounded float (abs unless signed) or '' when not numeric."""
    try:
        x = round(float(v), 2) if v not in (None, "") else ""
    except (ValueError, TypeError):
        return str(v) if v else ""
    return x if (x == "" or signed) else abs(x)


def _num(v):
    try:
        return float(str(v).replace(",", "")) if v not in (None, "") else None
    except (ValueError, TypeError):
        return None


def _fill_fields(r, order_lookup=None, manual_prices=None):
    """Display values for one fill, shared by the All Trades sheets.

    Trigger / limit are order-level fields, absent from Flex trades: they come
    from the captured order ledger (exact by permId, else best guess), and a
    price typed by hand into the previous report's All Trades sheet wins."""
    symbol   = r.get("symbol") or r.get("description") or ""
    raw_side = r.get("buySell", "").upper().strip()
    action   = ("BUY"  if raw_side in ("BUY", "BOT", "B") else
                "SELL" if raw_side in ("SELL", "SLD", "S") else raw_side)
    led_lim, led_trg = lookup_order_prices(order_lookup, symbol, action,
                                           r.get("orderType"), parse_trade_date(r),
                                           perm_id=r.get("permId") or r.get("ibOrderID"))
    trigger = (led_trg if led_trg is not None
               else r.get("triggerPrice") or r.get("auxPrice") or r.get("stopPrice"))
    limit   = (led_lim if led_lim is not None
               else r.get("limitPrice") or r.get("lmtPrice") or r.get("orderLimitPrice"))

    dt_str    = parse_trade_datetime(r)
    qty_str   = _fmt_abs(r.get("quantity"),   0)
    price_str = _fmt_abs(r.get("tradePrice"), 4)
    trigger_cell = _fmt_abs(trigger, 4)
    limit_cell   = _fmt_abs(limit,   4)
    if manual_prices:
        m_trg, m_lim = manual_prices.get(
            _trade_price_key(dt_str, symbol, action, qty_str, price_str), (None, None))
        if m_trg not in (None, ""):
            trigger_cell = m_trg
        if m_lim not in (None, ""):
            limit_cell = m_lim

    return {
        "dt": dt_str, "dt_gst": _to_gst(dt_str), "symbol": symbol,
        "name": r.get("description") or r.get("underlyingSymbol") or "",
        "action": action, "qty": qty_str, "price": price_str,
        "trigger": trigger_cell, "limit": limit_cell,
        "exchange": r.get("exchange", ""),
        "commission": _amt(r.get("ibCommission") or r.get("commission")),
        "realized": _amt(_row_realized_pnl(r), signed=True),
        "unrealized": ("" if r.get("_unrealized") is None
                       else _amt(r.get("_unrealized"), signed=True)),
        "account": r.get("accountId") or r.get("acctId") or r.get("account") or "",
    }


# ----------------------------------------------------------------------
# All Trades Detail: fills grouped by order. An entry order and its attached
# stop / exit orders share one Sr No (6, 6.1, 6.2 ...), numbered oldest first.
# ----------------------------------------------------------------------
DETAIL_HEADERS = ["Date & Time (UTC)", "Date & Time (GST)", "Sr No", "Contract", "Name",
                  "Action", "Quantity", "Price", "Order Type", "Trigger Price", "Offset",
                  "Limit Price", "SL %", "Exchange", "Commission", "Realized P&L",
                  "Unrealized P&L", "Account"]
DETAIL_WIDTHS  = [20, 20, 8, 12, 28, 8, 10, 12, 11, 12, 9, 12, 9, 12, 12, 13, 15, 13]


def build_detail_trade_rows(trade_rows, order_lookup=None, manual_prices=None):
    groups, n, out = {}, 0, []
    for idx, r in enumerate(sorted(trade_rows, key=_trade_sort_dt)):
        oid = r.get("permId") or r.get("ibOrderID")
        oid = str(oid) if oid not in (None, "", 0, "0") else None
        rec = order_lookup.by_perm.get(oid) if (order_lookup and oid) else None
        parent = str(rec["parentPermId"]) if rec and rec.get("parentPermId") else None

        gkey = parent or oid or f"fill-{idx}"
        if gkey not in groups:
            n += 1
            groups[gkey] = [n, 0]
            sr = str(n)
        else:
            groups[gkey][1] += 1
            sr = f"{groups[gkey][0]}.{groups[gkey][1]}"

        f = _fill_fields(r, order_lookup, manual_prices)
        trg, lim = _num(f["trigger"]), _num(f["limit"])
        offset = f"{abs(lim - trg):.4f}".rstrip("0").rstrip(".") if (trg and lim) else ""

        # SL %: on entry rows only - entry price vs its own stop-loss order.
        sl = ""
        if rec and not parent and oid:
            stops = [c for c in order_lookup.children.get(oid, [])
                     if _norm_type(c.get("orderType")).startswith("STP") and c.get("trigger")]
            entry = trg or lim
            if stops and entry:
                sl = f"{abs(entry - float(stops[0]['trigger'])) / entry * 100:.2f}%"

        out.append([
            f["dt"], f["dt_gst"], sr, f["symbol"], f["name"], f["action"],
            f["qty"], f["price"],
            r.get("orderType") or (rec or {}).get("orderType", ""),
            f["trigger"], offset, f["limit"], _sl_cell(f["action"], sl), f["exchange"],
            f["commission"], f["realized"], f["unrealized"], f["account"],
        ])
    out.reverse()                                   # newest first
    return out


def build_today_rows_from_flex(trade_rows):
    """Today Orders fallback: per-trade rows for today and yesterday only."""
    today = dt.date.today()
    return build_individual_trade_rows(trade_rows, {today, today - dt.timedelta(days=1)})


def _list_previous_reports():
    """All prior report files in the reports folder, newest first (by the date
    encoded in the filename, then modified time), excluding today's OUTPUT_FILE.
    Accepts the current "MIS_paper_" prefix and the legacy "MIS_"/"Reports_"
    prefixes."""
    try:
        candidates = []
        for fn in os.listdir(_REPORTS_DIR):
            if not fn.lower().endswith(".xlsx") or fn.startswith("~$"):
                continue
            path = os.path.join(_REPORTS_DIR, fn)
            if os.path.abspath(path) == os.path.abspath(OUTPUT_FILE):
                continue
            d = dt.date.min
            m = re.match(r"(?:MIS_paper_|MIS_|Reports_)(\d{2}[A-Za-z]{3}\d{4})", fn)
            if m:
                try:
                    d = dt.datetime.strptime(m.group(1), "%d%b%Y").date()
                except ValueError:
                    pass
            candidates.append((d, os.path.getmtime(path), path))
    except OSError:
        return []
    candidates.sort(key=lambda c: (c[0], c[1]), reverse=True)
    return [c[2] for c in candidates]


def _read_alltrades_prices(wb):
    """Build {trade_key: (trigger, limit)} from a workbook's 'All Trades' sheet.

    Columns are resolved by HEADER NAME (row 2), not fixed positions, so prices
    carry forward regardless of the file's column layout — older 14-column
    reports (single "Date & Time") and the current 15-column UTC/GST layout both
    read correctly, and future column changes won't silently break this."""
    prices = {}
    # Accept the current "paper_All Trades" tab and the legacy "All Trades" one
    # so prices still carry forward from reports written before the tab rename.
    sheet = next((nm for nm in ("paper_All Trades", "All Trades")
                  if nm in wb.sheetnames), None)
    if sheet is None:
        return prices
    rows = list(wb[sheet].iter_rows(min_row=2, values_only=True))
    if not rows:
        return prices
    header = [str(c) if c is not None else "" for c in rows[0]]

    def _col(*names):
        for n in names:
            if n in header:
                return header.index(n)
        return None

    c_dt  = _col("Date & Time (UTC)", "Date & Time")
    c_con = _col("Contract")
    c_act = _col("Action")
    c_qty = _col("Quantity")
    c_prc = _col("Price")
    c_trg = _col("Trigger Price")
    c_lim = _col("Limit Price")
    if None in (c_dt, c_con, c_act, c_qty, c_prc, c_trg, c_lim):
        return prices

    for row in rows[1:]:
        if not row:
            continue
        g = lambda i: row[i] if i < len(row) else None
        trigger, limit = g(c_trg), g(c_lim)
        if trigger in (None, "") and limit in (None, ""):
            continue
        key = _trade_price_key(g(c_dt) or "", g(c_con) or "", g(c_act) or "",
                               g(c_qty) or "", g(c_prc) or "")
        prices[key] = (
            str(trigger) if trigger not in (None, "") else None,
            str(limit)   if limit   not in (None, "") else None,
        )
    return prices


def load_previous_report():
    """Open the closest prior report (e.g. MIS_04Jun2026.xlsx) and pull forward
    the manual Trigger/Limit prices from the 'All Trades' sheet, which would
    otherwise be lost because every run writes a brand-new dated file. Prices are
    keyed by (datetime, contract, action, qty, price) so they re-attach to the
    same fill.

    Returns manual_prices: dict.
    """
    from openpyxl import load_workbook

    reports = _list_previous_reports()
    if not reports:
        return {}
    prev_path = reports[0]

    try:
        wb = load_workbook(prev_path, read_only=True, data_only=True)
    except Exception as e:                                  # noqa: BLE001
        print(f"  [Carry-forward] could not open {prev_path}: {e}", file=sys.stderr)
        return {}

    # Trigger/limit prices from the most recent report.
    manual_prices = _read_alltrades_prices(wb)

    wb.close()

    # Self-heal: if the latest report carries no trigger/limit prices at all
    # (e.g. a column-layout change once broke the chain), fall back to the most
    # recent older report that still has them so the prices aren't lost.
    if not manual_prices:
        for older in reports[1:]:
            try:
                w2 = load_workbook(older, read_only=True, data_only=True)
            except Exception:                              # noqa: BLE001
                continue
            mp = _read_alltrades_prices(w2)
            w2.close()
            if mp:
                manual_prices = mp
                print(f"  [Carry-forward] recovered {len(mp)} trigger/limit "
                      f"price(s) from {os.path.basename(older)}.")
                break
    print(f"  [Carry-forward] {len(manual_prices)} trigger/limit price(s) "
          f"from {os.path.basename(prev_path)}.")
    return manual_prices


def _exec_to_flex_row(ex):
    """Map a live TWS execution into a Flex-style trade dict so it flows through
    the same aggregation / realized-PnL pipeline as Flex trades. Commission is
    stored as a negative cost to match the Flex 'ibCommission' convention, and
    fifoPnlRealized is left blank so the FIFO engine computes it."""
    comm = ex.get("commission")
    return {
        # The live execId is the same string Flex later reports as ibExecID, so
        # carrying it here lets a live fill and its Flex republish share one key.
        "ibExecID":        ex.get("execId", ""),
        "symbol":          ex.get("contract", ""),
        "description":     ex.get("contract", ""),
        "buySell":         ex.get("action", ""),
        "quantity":        ex.get("quantity", 0),
        "tradePrice":      ex.get("price", 0),
        "commission":      (-abs(_flt(comm)) if comm is not None else None),
        "exchange":        ex.get("exchange", ""),
        "accountId":       ex.get("account", ""),
        "dateTime":        " ".join(str(ex.get("time", "")).split()),
        "orderType":       "",
        "fifoPnlRealized": "",
        "permId":          ex.get("permId") or "",
        "conid":           str(ex.get("conId") or ""),
    }


# ----------------------------------------------------------------------
# STEP 1: request the statement -> reference code
# ----------------------------------------------------------------------
def send_request():
    r = requests.get(SEND_URL, params={"t": TOKEN, "q": QUERY_ID, "v": VERSION}, timeout=30)
    r.raise_for_status()
    root = ET.fromstring(r.content)
    status = root.findtext("Status")
    if status != "Success":
        code = root.findtext("ErrorCode")
        msg  = root.findtext("ErrorMessage")
        raise RuntimeError(f"SendRequest failed [{code}]: {msg}")
    ref = root.findtext("ReferenceCode")
    print(f"Reference code: {ref}")
    return ref


# ----------------------------------------------------------------------
# STEP 2: poll for the statement XML
# ----------------------------------------------------------------------
def get_statement(ref):
    for attempt in range(1, MAX_POLL_ATTEMPTS + 1):
        r = requests.get(GET_URL, params={"t": TOKEN, "q": ref, "v": VERSION}, timeout=60)
        r.raise_for_status()
        text = r.text
        if "<FlexQueryResponse" in text:
            print("Statement retrieved.")
            return text
        try:
            root = ET.fromstring(r.content)
            code = root.findtext("ErrorCode")
            msg  = root.findtext("ErrorMessage")
            if code and code not in ("1019",):
                raise RuntimeError(f"GetStatement failed [{code}]: {msg}")
            print(f"  not ready yet (attempt {attempt}/{MAX_POLL_ATTEMPTS}), waiting...")
        except ET.ParseError:
            pass
        time.sleep(POLL_WAIT_SECONDS)
    raise RuntimeError("Statement not ready after maximum retries.")


# ----------------------------------------------------------------------
# STEP 3: parse trades + account ID
# ----------------------------------------------------------------------
def parse_trades(xml_text):
    root   = ET.fromstring(xml_text)
    trades = root.findall(".//Trade")
    if not trades:
        trades = root.findall(".//Order")
    return [dict(t.attrib) for t in trades]


def parse_account_id(xml_text):
    root = ET.fromstring(xml_text)
    stmt = root.find(".//FlexStatement")
    if stmt is not None:
        return stmt.get("accountId", "N/A")
    return "N/A"


def parse_trade_date(row):
    """Extract a date from common IBKR datetime attributes."""
    raw = row.get("dateTime") or row.get("tradeDate") or row.get("reportDate") or ""
    raw = raw.strip()
    if not raw:
        return None
    raw = raw.replace(",", " ").replace(";", " ")
    datepart = raw.split()[0]
    for fmt in ("%Y%m%d", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return dt.datetime.strptime(datepart, fmt).date()
        except ValueError:
            continue
    return None


def parse_trade_datetime(row):
    """Extract a full datetime string from IBKR attributes, formatted for display.

    IBKR Flex sends dateTime as 'DD/MM/YYYY;HHMMSS' (e.g. '20/05/2026;080617'),
    but other configs use 'YYYYMMDD;HHMMSS' or ISO. We normalise separators
    then parse the date and time parts independently.
    """
    raw = (row.get("dateTime") or row.get("tradeDate") or "").strip()
    if not raw:
        return ""
    norm  = raw.replace(";", " ").replace(",", " ")
    parts = norm.split()
    datepart = parts[0]
    timepart = parts[1] if len(parts) > 1 else ""

    d = None
    for fmt in ("%Y%m%d", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            d = dt.datetime.strptime(datepart, fmt)
            break
        except ValueError:
            continue
    if d is None:
        return raw

    if timepart:
        for tfmt in ("%H%M%S", "%H:%M:%S"):
            try:
                t = dt.datetime.strptime(timepart, tfmt).time()
                return dt.datetime.combine(d.date(), t).strftime("%d-%b-%Y %H:%M:%S")
            except ValueError:
                continue
    return d.strftime("%d-%b-%Y")


# ----------------------------------------------------------------------
# STEP 4: aggregate raw trades into per-contract summary rows
# ----------------------------------------------------------------------
HEADERS = [
    "First Trade Date (UTC)", "First Trade Date (GST)",
    "Last Trade Date (UTC)",  "Last Trade Date (GST)",
    "Contract", "Name", "Buys", "Sells", "Net",
    "Avg (bought)", "Avg (sold)",
    "Total (bought)", "Total (sold)",
    "Exchange List",
    "Commission",
    "Realized P&L", "Unrealized P&L",
]

# Trade Summary columns shaded green (positive) / red (negative).
# Commission is intentionally excluded — it is shown with no colour.
COLOR_COLUMNS_SUMMARY = ["Realized P&L", "Unrealized P&L"]

# Money columns formatted as 1,234.00 on the aggregate sheets.
AMOUNT_COLUMNS_SUMMARY = ["Avg (bought)", "Avg (sold)", "Total (bought)",
                          "Total (sold)", "Commission", "Realized P&L", "Unrealized P&L"]


def _flt(val):
    try:
        return float(val) if val not in (None, "") else 0.0
    except (ValueError, TypeError):
        return 0.0


def _fmt_limit(val):
    """Format a risk-limit value as 1,234.00 when numeric, else show it as-is
    (e.g. the placeholder 'NA')."""
    try:
        return f"{float(val):,.2f}"
    except (ValueError, TypeError):
        return str(val)


def aggregate(trade_rows, portfolio=None):
    """Per-contract summary. Realized P&L sums the given fills; Unrealized P&L
    is IB's live figure for the contract's open position (0 when flat, blank
    when the Gateway gave no portfolio)."""
    data = defaultdict(lambda: {
        "buy_qty": 0.0, "sell_qty": 0.0,
        "buy_value": 0.0, "sell_value": 0.0,
        "commission": 0.0, "pnl": 0.0, "unrealized": 0.0,
        "exchanges": set(),
        "dates": [],
        "name": "",
    })

    for r in trade_rows:
        symbol     = r.get("symbol") or r.get("description") or "UNKNOWN"
        name       = r.get("description") or r.get("underlyingSymbol") or ""
        side       = r.get("buySell", "").upper().strip()
        qty        = abs(_flt(r.get("quantity", 0)))
        price      = abs(_flt(r.get("tradePrice", 0)))
        money      = abs(_flt(r.get("tradeMoney", 0))) or qty * price
        # IBKR Flex uses 'ibCommission'; the TWS API uses 'commission'.
        commission = _flt(r.get("ibCommission") or r.get("commission") or 0)
        pnl        = _row_realized_pnl(r)
        exchange   = r.get("exchange", "")

        c = data[symbol]
        if side in ("BUY", "B"):
            c["buy_qty"]   += qty
            c["buy_value"] += money
        elif side in ("SELL", "S"):
            c["sell_qty"]   += qty
            c["sell_value"] += money
        c["commission"] += commission
        c["pnl"]        += pnl
        if exchange:
            c["exchanges"].add(exchange)
        if name and not c["name"]:
            c["name"] = name
        d = parse_trade_date(r)
        if d:
            c["dates"].append((d, parse_trade_datetime(r)))

    result = []
    for symbol, c in data.items():
        avg_b     = c["buy_value"]  / c["buy_qty"]  if c["buy_qty"]  else 0.0
        avg_s     = c["sell_value"] / c["sell_qty"] if c["sell_qty"] else 0.0
        datetimes = sorted(c["dates"], key=lambda x: x[0])
        result.append({
            "Contract":             symbol,
            "Name":                 c["name"],
            "Buys":                 round(c["buy_qty"], 4),
            "Sells":                round(c["sell_qty"], 4),
            "Net":                  round(c["buy_qty"] - c["sell_qty"], 4),
            "Avg (bought)":         round(avg_b, 6),
            "Avg (sold)":           round(avg_s, 6),
            "Total (bought)":       round(c["buy_value"], 2),
            "Total (sold)":         round(c["sell_value"], 2),
            "Exchange List":        ", ".join(sorted(c["exchanges"])),
            "Commission":           round(c["commission"], 2),
            "Realized P&L":         round(c["pnl"], 2),
            "Unrealized P&L":       ("" if portfolio is None else
                                     round(_flt((portfolio.get(symbol) or {}).get("unrealized")), 2)),
            "First Trade Date (UTC)": datetimes[0][1]  if datetimes else "",
            "First Trade Date (GST)": _to_gst(datetimes[0][1])  if datetimes else "",
            "Last Trade Date (UTC)":  datetimes[-1][1] if datetimes else "",
            "Last Trade Date (GST)":  _to_gst(datetimes[-1][1]) if datetimes else "",
        })
    result.sort(key=lambda x: x["Contract"])
    return result


def _trade_sort_dt(row):
    """A real datetime for chronological (FIFO) ordering of fills.
    Falls back to datetime.min when the row carries no parseable date."""
    d = parse_trade_date(row)
    if d is None:
        return dt.datetime.min
    raw   = (row.get("dateTime") or row.get("tradeDate") or "").strip()
    parts = raw.replace(";", " ").replace(",", " ").split()
    timepart = parts[1] if len(parts) > 1 else ""
    for tfmt in ("%H%M%S", "%H:%M:%S"):
        try:
            t = dt.datetime.strptime(timepart, tfmt).time()
            return dt.datetime.combine(d, t)
        except ValueError:
            continue
    return dt.datetime.combine(d, dt.time.min)


def _fill_cpu_mult(r):
    """(commission per unit, contract multiplier) for one fill."""
    qty  = abs(_flt(r.get("quantity", 0)))
    comm = abs(_flt(r.get("ibCommission") or r.get("commission") or 0))
    mult = _flt(r.get("multiplier")) or 1.0
    return (comm / qty if qty else 0.0), mult


def compute_realized_pnl(trade_rows):
    """FIFO realized P&L per fill, net of commissions.

    Runs over the FULL stored history (not just the report window) so a
    position opened before the window start is closed against its real cost.
    Each closing fill is matched against earlier opposite-side lots, first in
    first out; realized = price difference x matched qty x multiplier, minus
    the commission share of both the opening and the closing fill. It is
    stored on the closing fill as '_calcRealizedPnl'. IBKR's own Flex figure
    (fifoPnlRealized, also net of commission) is preferred when it is given.

    Also stores '_openQty' on each fill: the part of it that is still open
    (signed, + long / - short), used for per-fill Unrealized P&L."""
    from collections import deque

    lots_by_symbol = defaultdict(deque)   # symbol -> [qty, price, sign, cpu, mult, row]
    for r in sorted(trade_rows, key=_trade_sort_dt):
        r["_openQty"] = 0.0
        symbol = r.get("symbol") or r.get("description") or "UNKNOWN"
        side   = r.get("buySell", "").upper().strip()
        qty    = abs(_flt(r.get("quantity", 0)))
        price  = abs(_flt(r.get("tradePrice", 0)))
        sign   = (1 if side in ("BUY", "BOT", "B") else
                  -1 if side in ("SELL", "SLD", "S") else 0)
        if qty == 0 or sign == 0:
            r["_calcRealizedPnl"] = 0.0
            continue
        cpu, mult = _fill_cpu_mult(r)

        lots      = lots_by_symbol[symbol]
        realized  = 0.0
        remaining = qty
        while remaining > 1e-9 and lots and lots[0][2] == -sign:
            lot     = lots[0]
            matched = min(remaining, lot[0])
            diff    = (price - lot[1]) if lot[2] == 1 else (lot[1] - price)
            realized += diff * matched * lot[4] - matched * (lot[3] + cpu)
            lot[0]    -= matched
            remaining -= matched
            lot[5]["_openQty"] = lot[0] * lot[2]
            if lot[0] <= 1e-9:
                lots.popleft()
        if remaining > 1e-9:
            lots.append([remaining, price, sign, cpu, mult, r])
            r["_openQty"] = remaining * sign
        r["_calcRealizedPnl"] = round(realized, 2)


def annotate_unrealized(trade_rows, portfolio):
    """Per-fill Unrealized P&L for the still-open part of each fill, valued at
    IB's current market price: open qty x (market - fill price) x multiplier,
    minus that part's commission. Blank when no live price is available."""
    for r in trade_rows:
        r["_unrealized"] = None
        oq = _flt(r.get("_openQty"))
        if not oq or portfolio is None:
            continue
        pos = portfolio.get(r.get("symbol") or "")
        mkt = _flt(pos.get("marketPrice")) if pos else 0.0
        if not mkt:
            continue
        cpu, mult = _fill_cpu_mult(r)
        price = abs(_flt(r.get("tradePrice", 0)))
        r["_unrealized"] = round(oq * (mkt - price) * mult - abs(oq) * cpu, 2)


def _row_realized_pnl(row):
    """Realized P&L for one fill: the official Flex 'fifoPnlRealized' when the
    feed reports a real (non-zero) figure, otherwise our FIFO-computed
    '_calcRealizedPnl'. Paper accounts send 'fifoPnlRealized=0' for every fill,
    so a zero/blank Flex value is treated as "not provided" and we fall back to
    the calculated value rather than masking it with 0."""
    flex = _flt(row.get("fifoPnlRealized"))
    if flex != 0:
        return flex
    return _flt(row.get("_calcRealizedPnl", 0))


# ----------------------------------------------------------------------
# STEP 5: write Excel
# ----------------------------------------------------------------------

# Style constants — clean corporate palette: navy headers, neutral text
_NAVY        = PatternFill("solid", fgColor="1F3864")   # dark navy (title + header)
_TITLE_FONT  = Font(bold=True, size=13, color="FFFFFF", name="Calibri")
_HDR_FONT    = Font(bold=True, size=10, color="FFFFFF", name="Calibri")
_DATA_FONT   = Font(size=10, color="262626", name="Calibri")   # near-black body text
_LABEL_FONT  = Font(bold=True, size=10, color="1F3864", name="Calibri")
_CENTER      = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT        = Alignment(horizontal="left",   vertical="center")
_RIGHT       = Alignment(horizontal="right",  vertical="center")

# PnL cell highlights: green for profit, faint red for loss.
_PNL_POS_FILL = PatternFill("solid", fgColor="C6EFCE")   # light green (positive)
_PNL_NEG_FILL = PatternFill("solid", fgColor="FFC7CE")   # faint red (negative)

# Thin grey gridline border applied around every cell in a table.
_GRID_SIDE   = Side(style="thin", color="BFBFBF")
_CELL_BORDER = Border(left=_GRID_SIDE, right=_GRID_SIDE, top=_GRID_SIDE, bottom=_GRID_SIDE)

# Uniform amount/money display across every sheet, e.g. 1,234.00 (and -1,234.00).
_AMOUNT_FMT = "#,##0.00"


def _apply_amount_format(ws, col, start_row=3):
    """Apply the 1,234.00 number format to numeric cells in a column."""
    for r in range(start_row, ws.max_row + 1):
        cell = ws.cell(row=r, column=col)
        if isinstance(cell.value, (int, float)):
            cell.number_format = _AMOUNT_FMT

# Green (Excel selection green) border drawn around each sheet's title banner.
_TITLE_SIDE   = Side(style="medium", color="00B050")
_TITLE_BORDER = Border(left=_TITLE_SIDE, right=_TITLE_SIDE, top=_TITLE_SIDE, bottom=_TITLE_SIDE)


def _apply_cell_borders(ws, n_cols, start_row=2):
    """Draw a thin border around every cell from start_row..max_row, cols 1..n_cols."""
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row, max_col=n_cols):
        for cell in row:
            cell.border = _CELL_BORDER


def _color_pnl_cell(cell):
    """Shade a cell green if its numeric value is positive, faint red if negative."""
    try:
        v = float(str(cell.value).replace(",", ""))
    except (ValueError, TypeError, AttributeError):
        return
    if v > 0:
        cell.fill = _PNL_POS_FILL
    elif v < 0:
        cell.fill = _PNL_NEG_FILL

# Every worksheet tab name carries this prefix (paper-trading report).
_TAB_PREFIX = "paper_"


def _tab(name):
    """Worksheet tab name: the internal sheet key prefixed with 'paper_'."""
    return f"{_TAB_PREFIX}{name}"


# Sheet display titles
def _sheet_titles():
    frm = DATA_FROM.strftime("%d-%b-%Y")
    return {
        "Dashboard":     "ACCOUNT DASHBOARD",
        "Pending Order": "PENDING ORDERS",
        "All Trades":    f"ALL TRADES  (From {frm})",
        "All Trades Detail": f"ALL TRADES DETAIL  (From {frm})",
        "Open Position": f"OPEN POSITIONS  (From {frm}, live P&L)",
        "Trade Summary": f"TRADE SUMMARY  (From {frm})",
    }


def _write_title_row(ws, text, n_cols):
    """Row 1: merged dark-navy title spanning all columns, framed by a green border."""
    end_col = get_column_letter(n_cols)
    ws.merge_cells(f"A1:{end_col}1")
    cell = ws["A1"]
    cell.value     = text
    cell.font      = _TITLE_FONT
    cell.fill      = _NAVY
    cell.alignment = _CENTER
    ws.row_dimensions[1].height = 24
    # Green frame around the whole merged title banner. For a merged cell the
    # border must be set on every underlying cell so all four outer edges show.
    for col in range(1, n_cols + 1):
        ws.cell(row=1, column=col).border = _TITLE_BORDER


def _write_header_row(ws, headers, row=2):
    """Row 2: per-column headers in dark navy."""
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font      = _HDR_FONT
        cell.fill      = _NAVY
        cell.alignment = _CENTER
    ws.row_dimensions[row].height = 30


def _style_data_rows(ws, start_row, n_cols):
    """Apply data font and alignment to all data rows."""
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row, max_col=n_cols):
        for cell in row:
            cell.font      = _DATA_FONT
            cell.alignment = _CENTER


def _apply_autofilter(ws, n_cols, header_row=2):
    """Enable Excel column filters on the header row across all columns."""
    last_col = get_column_letter(n_cols)
    last_row = max(ws.max_row, header_row)
    ws.auto_filter.ref = f"A{header_row}:{last_col}{last_row}"


def write_excel(rows, pending_rows, trade_rows, account_id, account_data, order_lookup=None,
                portfolio=None):
    today     = dt.date.today()
    yesterday = today - dt.timedelta(days=1)
    cutoff_7  = today - dt.timedelta(days=7)

    # Realized P&L (FIFO) over the FULL history first, so positions opened
    # before the report window are closed against their real cost.
    all_rows = rows
    compute_realized_pnl(all_rows)
    annotate_unrealized(all_rows, portfolio)

    # ── Global cutoff: keep only trades on/after DATA_FROM ──
    rows = [r for r in rows if (parse_trade_date(r) or dt.date.min) >= DATA_FROM]
    print(f"  [Filter] {len(rows)} trade(s) on/after {DATA_FROM:%d-%b-%Y}.")

    # Calculate FIFO realized P&L per fill (annotates each row in place) so the
    # All Trades, Trade Summary and Dashboard all reflect a computed P&L even
    # when IBKR's Flex feed omits it (e.g. on paper accounts).

    # Today's set mirrors the Today Orders sheet: today's fills, or yesterday's
    # as a fallback when IBKR's Flex batch hasn't published today's yet.
    rows_today, rows_7 = [], []
    for row in rows:
        d = parse_trade_date(row)
        if d is None:
            continue
        if d in (today, yesterday):
            rows_today.append(row)
        if d >= cutoff_7:
            rows_7.append(row)

    agg_today = aggregate(rows_today)
    agg_7     = aggregate(rows_7)
    agg_all   = aggregate(rows, portfolio)  # all trades since DATA_FROM
    pending_rows = add_pnl_to_pending(pending_rows, agg_all, portfolio)

    # Carry forward the manual trigger/limit prices (All Trades) from the
    # previous report.
    manual_prices = load_previous_report()

    # All Trades: every individual trade since DATA_FROM (ledger trigger/limit,
    # with manually-entered prices from the previous report taking precedence).
    all_trade_rows = build_individual_trade_rows(rows, order_lookup=order_lookup,
                                                 manual_prices=manual_prices)
    print(f"  [All Trades] {len(all_trade_rows)} individual trade(s) since {DATA_FROM:%d-%b-%Y}.")
    detail_rows = build_detail_trade_rows(rows, order_lookup=order_lookup,
                                          manual_prices=manual_prices)

    # Sheet sequence: Index, Dashboard, Pending Order, Open Position,
    # All Trades, All Trades Detail, Trade Summary.
    wb = Workbook()
    _fill_index_sheet(wb.active,
                      ["Dashboard", "Pending Order", "Open Position",
                       "All Trades", "All Trades Detail", "Trade Summary"],
                      account_id)
    _fill_dashboard_sheet(wb.create_sheet(), account_id, account_data, pending_rows,
                          agg_7, agg_today, agg_all, len(all_trade_rows), portfolio)
    _fill_pending_sheet(wb.create_sheet(),            pending_rows)
    _fill_running_positions_sheet(wb.create_sheet(),  rows, portfolio)
    _fill_trade_list_sheet(wb.create_sheet(), "All Trades", all_trade_rows)
    _fill_table_sheet(wb.create_sheet(), "All Trades Detail", DETAIL_HEADERS,
                      DETAIL_WIDTHS, detail_rows,
                      amount_cols=("Commission", "Realized P&L", "Unrealized P&L"),
                      color_cols=("Realized P&L", "Unrealized P&L"))
    _fill_sheet(wb.create_sheet(), "Trade Summary",   agg_all)
    os.makedirs(_REPORTS_DIR, exist_ok=True)
    wb.save(OUTPUT_FILE)
    print(f"Wrote {OUTPUT_FILE}: {len(pending_rows)} pending, "
          f"{len(all_trade_rows)} all trades, "
          f"{len(agg_today)} contracts (today), "
          f"{len(agg_7)} contracts (7d), {len(agg_all)} contracts (summary)")


# ── Index sheet (links to every other sheet) ──────────────────────────
def _ordinal(n):
    if 11 <= (n % 100) <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def _fill_index_sheet(ws, sheet_titles, account_id="N/A"):
    ws.title = _tab("Index")

    title_fill = PatternFill("solid", fgColor="DDEBF7")   # light blue
    head_fill  = PatternFill("solid", fgColor="DDD9C4")   # beige
    data_fill  = PatternFill("solid", fgColor="DDEBF7")   # light blue
    thin       = Side(style="thin", color="1F3864")
    border     = Border(left=thin, right=thin, top=thin, bottom=thin)
    link_font  = Font(name="Calibri", size=11, bold=True, color="0563C1", underline="single")
    head_font  = Font(name="Calibri", size=11, bold=True, color="1F3864")
    title_font = Font(name="Calibri", size=12, bold=True, color="1F3864")
    label_font = Font(name="Calibri", size=11, bold=True, color="1F3864")
    value_font = Font(name="Calibri", size=11, color="1F3864")

    today = dt.date.today()
    report_title = f"IBKR - MIS Report as on {_ordinal(today.day)} {today:%B %Y}"

    # Row 3: report title (merged B:C)
    ws.merge_cells("B3:C3")
    t = ws["B3"]
    t.value, t.font, t.fill = report_title, title_font, title_fill
    t.alignment = Alignment(horizontal="center", vertical="center")

    # Row 4: "Index" (merged B:C)
    ws.merge_cells("B4:C4")
    ix = ws["B4"]
    ix.value, ix.font, ix.fill = "Index", head_font, title_fill
    ix.alignment = Alignment(horizontal="center", vertical="center")

    # Row 5: column headers
    for col, label in ((2, "SI"), (3, "Description")):
        c = ws.cell(row=5, column=col, value=label)
        c.font, c.fill = head_font, head_fill
        c.alignment = Alignment(horizontal="center", vertical="center")

    # Rows 6+: one linked row per sheet
    for i, title in enumerate(sheet_titles, start=1):
        r = 5 + i
        sl = ws.cell(row=r, column=2, value=i)
        sl.fill = data_fill
        sl.alignment = Alignment(horizontal="center", vertical="center")
        sl.font = Font(name="Calibri", size=11, color="1F3864")

        link = ws.cell(row=r, column=3, value=_tab(title))
        link.hyperlink  = f"#'{_tab(title)}'!A1"
        link.font       = link_font
        link.fill       = data_fill
        link.alignment  = Alignment(horizontal="center", vertical="center")

    # Borders around the whole table (rows 3..last, cols B:C)
    last_row = 5 + len(sheet_titles)
    for row in ws.iter_rows(min_row=3, max_row=last_row, min_col=2, max_col=3):
        for cell in row:
            cell.border = border

    # Green frame around the report title banner (B3:C3), matching other sheets.
    ws["B3"].border = _TITLE_BORDER
    ws["C3"].border = _TITLE_BORDER

    # ── Report Details box (cols E:F), alongside the index table ──────
    details = [
        ("Account ID",         account_id),
        ("Report Date",        today.strftime("%d-%b-%Y")),
        ("Data From (cutoff)", DATA_FROM.strftime("%d-%b-%Y")),
    ]
    ws.merge_cells("E4:F4")
    dt_title = ws["E4"]
    dt_title.value, dt_title.font, dt_title.fill = "Report Details", head_font, title_fill
    dt_title.alignment = Alignment(horizontal="center", vertical="center")
    for i, (label, value) in enumerate(details, start=1):
        r = 4 + i
        lc = ws.cell(row=r, column=5, value=label)
        lc.font, lc.fill = label_font, data_fill
        lc.alignment = Alignment(horizontal="left", vertical="center")
        vc = ws.cell(row=r, column=6, value=value)
        vc.font, vc.fill = value_font, data_fill
        vc.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=4, max_row=4 + len(details), min_col=5, max_col=6):
        for cell in row:
            cell.border = border

    # ── Login Details box (cols E:F), below the Report Details box ─────
    # Report Details fills rows 4..(4+len(details)); leave one blank row, then
    # start the credentials box beneath it.
    cred_title_row = 4 + len(details) + 2          # one-row gap after Report Details
    credentials = [
        ("Username", "gautam575"),
        ("Password", "algoggt1))"),
    ]
    ws.merge_cells(start_row=cred_title_row, start_column=5,
                   end_row=cred_title_row, end_column=6)
    ct = ws.cell(row=cred_title_row, column=5)
    ct.value, ct.font, ct.fill = "Login Details (Paper Account)", head_font, title_fill
    ct.alignment = Alignment(horizontal="center", vertical="center")
    for i, (label, value) in enumerate(credentials, start=1):
        r = cred_title_row + i
        lc = ws.cell(row=r, column=5, value=label)
        lc.font, lc.fill = label_font, data_fill
        lc.alignment = Alignment(horizontal="left", vertical="center")
        vc = ws.cell(row=r, column=6, value=value)
        vc.font, vc.fill = value_font, data_fill
        vc.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=cred_title_row,
                            max_row=cred_title_row + len(credentials),
                            min_col=5, max_col=6):
        for cell in row:
            cell.border = border

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 34
    ws.column_dimensions["E"].width = 20
    ws.column_dimensions["F"].width = 16
    ws.sheet_view.showGridLines = False


# ── Dashboard sheet ────────────────────────────────────────────────────
def _fill_dashboard_sheet(ws, account_id, account_data, pending_rows,
                          agg_7, agg_today, agg_all, num_all_trades, portfolio=None):
    ws.title  = _tab("Dashboard")
    now_ist   = dt.datetime.now(_IST)
    n_cols    = 2   # Parameter | Value

    _write_title_row(ws, _sheet_titles()["Dashboard"], n_cols)
    _write_header_row(ws, ["Parameter", "Value"], row=2)

    # ── Report info ───────────────────────────────────────────────────
    info = [
        ("Account ID",         account_id,                     ""),
        ("Report Date",        now_ist.strftime("%d-%b-%Y"),    ""),
        ("Data From (cutoff)", DATA_FROM.strftime("%d-%b-%Y"),  ""),
    ]

    # ── All-trades summary (everything since DATA_FROM) ───────────────
    at_buys   = sum(r["Total (bought)"]       for r in agg_all)
    at_sells  = sum(r["Total (sold)"]         for r in agg_all)
    at_comm   = sum(r["Commission"]           for r in agg_all)
    at_pnl    = sum(r["Realized P&L"]         for r in agg_all)
    # Unrealized: every open position right now (live from the Gateway).
    at_unreal = (None if portfolio is None else
                 sum(_flt(p.get("unrealized")) for p in portfolio.values()))

    summary_all = [
        ("Total Trades (fills)", num_all_trades,            ""),
        ("Contracts Traded",     len(agg_all),              ""),
        ("Total Bought Value",   f"{at_buys:,.2f}",          ""),
        ("Total Sold Value",     f"{at_sells:,.2f}",         ""),
        ("Commission",           f"{at_comm:,.2f}",          ""),
        ("Realized P&L",         f"{at_pnl:,.2f}",           ""),
        ("Unrealized P&L",       ("N/A (Gateway not connected)" if at_unreal is None
                                  else f"{at_unreal:,.2f}"),  ""),
        ("Open Positions",       ("N/A" if portfolio is None else len(portfolio)), ""),
        ("Max Exposure",         _fmt_limit(MAX_EXPOSURE),    ""),
        ("Daily Max Loss",       _fmt_limit(DAILY_MAX_LOSS),  ""),
        ("Loss due to bugs",     f"{LOSS_DUE_TO_BUGS:,.2f}", ""),
    ]

    # Dashboard value cells shaded green (positive) / red (negative).
    # Commission is intentionally excluded — it is shown with no colour.
    color_labels = {"Loss due to bugs", "Unrealized P&L", "Realized P&L"}

    # ── Assemble all sections with dividers ───────────────────────────
    sections = [
        ("REPORT INFO",                          info),
        (f"ALL TRADES SUMMARY (From {DATA_FROM:%d-%b-%Y})", summary_all),
    ]

    r_idx = 3
    for section_title, section_rows in sections:
        # Blank gap before each section (except first)
        if r_idx > 3:
            ws.row_dimensions[r_idx].height = 6
            r_idx += 1

        # Section sub-heading row
        for col in range(1, n_cols + 1):
            cell = ws.cell(row=r_idx, column=col,
                           value=section_title if col == 1 else "")
            cell.font      = Font(bold=True, size=10, color="FFFFFF", name="Calibri")
            cell.fill      = _NAVY
            cell.alignment = _LEFT
        ws.row_dimensions[r_idx].height = 18
        r_idx += 1

        # Data rows (Currency column dropped — values are USD as shown in TWS)
        for label, value, _currency in section_rows:
            ca = ws.cell(row=r_idx, column=1, value=label)
            cb = ws.cell(row=r_idx, column=2, value=value)
            ca.font = _LABEL_FONT
            cb.font = _DATA_FONT
            cb.alignment = _RIGHT
            if label in color_labels:
                _color_pnl_cell(cb)
            ws.row_dimensions[r_idx].height = 16
            r_idx += 1

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 22
    ws.freeze_panes = "A3"


# ── Per-trade sheets (All Trades) ─────────────────────────────────────
def _fill_trade_list_sheet(ws, sheet_key, rows):
    ws.title  = _tab(sheet_key)
    n_cols    = len(TODAY_HEADERS)

    _write_title_row(ws, _sheet_titles()[sheet_key], n_cols)
    _write_header_row(ws, TODAY_HEADERS, row=2)

    for row in rows:
        ws.append(row)

    _style_data_rows(ws, start_row=3, n_cols=n_cols)

    # Amount columns: uniform 1,234.00 number format.
    for h in ("Commission", "Realized P&L", "Unrealized P&L"):
        _apply_amount_format(ws, TODAY_HEADERS.index(h) + 1)

    # Shade the Commission / Realized PnL cells green (positive) / red (negative).
    for h in COLOR_COLUMNS_TRADES:
        col = TODAY_HEADERS.index(h) + 1
        for r in range(3, ws.max_row + 1):
            _color_pnl_cell(ws.cell(row=r, column=col))

    for col, width in enumerate(TODAY_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width

    _apply_cell_borders(ws, n_cols, start_row=2)
    _apply_autofilter(ws, n_cols)
    ws.freeze_panes = "A3"


# ── Generic table sheet (All Trades Detail) ───────────────────────────
def _fill_table_sheet(ws, key, headers, widths, rows, amount_cols=(), color_cols=()):
    ws.title = _tab(key)
    n_cols   = len(headers)

    _write_title_row(ws, _sheet_titles()[key], n_cols)
    _write_header_row(ws, headers, row=2)
    for row in rows:
        ws.append(row)
    _style_data_rows(ws, start_row=3, n_cols=n_cols)

    for h in amount_cols:
        _apply_amount_format(ws, headers.index(h) + 1)
    for h in color_cols:
        col = headers.index(h) + 1
        for r in range(3, ws.max_row + 1):
            _color_pnl_cell(ws.cell(row=r, column=col))

    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    _apply_cell_borders(ws, n_cols, start_row=2)
    _apply_autofilter(ws, n_cols)
    ws.freeze_panes = "A3"


# ── Pending Order sheet ────────────────────────────────────────────────
def _fill_pending_sheet(ws, rows):
    ws.title  = _tab("Pending Order")
    n_cols    = len(PENDING_HEADERS)

    _write_title_row(ws, _sheet_titles()["Pending Order"], n_cols)
    _write_header_row(ws, PENDING_HEADERS, row=2)

    for row in rows:
        ws.append(row)

    _style_data_rows(ws, start_row=3, n_cols=n_cols)
    for h in ("Realized P&L", "Unrealized P&L"):
        col = PENDING_HEADERS.index(h) + 1
        _apply_amount_format(ws, col)
        for r in range(3, ws.max_row + 1):
            _color_pnl_cell(ws.cell(row=r, column=col))

    for col, width in enumerate(PENDING_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width

    _apply_autofilter(ws, n_cols)
    ws.freeze_panes = "A3"


# ── Open Position sheet ───────────────────────────────────────────────
def _fill_running_positions_sheet(ws, trade_rows, portfolio=None):
    # Figures (dates, buys/sells, realized P&L) cover the report window only;
    # which positions are open, their size and unrealized P&L come live from IB.
    agg     = aggregate(trade_rows, portfolio)
    if portfolio is None:
        # Gateway not connected: fall back to the net quantity from history.
        running = [r for r in agg if r["Net"] != 0]
    else:
        # IB's live portfolio is the authority on what is open (FX conversions,
        # for example, are cash balances to IB, not positions).
        running = [dict(r, Net=portfolio[r["Contract"]]["position"])
                   for r in agg if r["Contract"] in portfolio]
        known = {r["Contract"] for r in running}
        for sym, p in portfolio.items():            # open, but not traded in the window
            if sym not in known:
                running.append({h: "" for h in HEADERS} | {
                    "Contract": sym, "Net": p["position"], "Realized P&L": 0.0,
                    "Unrealized P&L": round(_flt(p.get("unrealized")), 2)})
        running.sort(key=lambda r: r["Contract"])
    _fill_sheet(ws, "Open Position", running)


# ── Generic trade aggregation sheet ──────────────────────────────────
def _fill_sheet(ws, title, agg_rows):
    ws.title = _tab(title)
    n_cols   = len(HEADERS)

    _write_title_row(ws, _sheet_titles().get(title, title.upper()), n_cols)
    _write_header_row(ws, HEADERS, row=2)

    for row in agg_rows:
        ws.append([row.get(h, "") for h in HEADERS])

    _style_data_rows(ws, start_row=3, n_cols=n_cols)

    # Amount columns: uniform 1,234.00 number format.
    for h in AMOUNT_COLUMNS_SUMMARY:
        if h in HEADERS:
            _apply_amount_format(ws, HEADERS.index(h) + 1)

    # Shade money columns green (positive) / faint red (negative).
    for h in COLOR_COLUMNS_SUMMARY:
        col = HEADERS.index(h) + 1
        for r in range(3, ws.max_row + 1):
            _color_pnl_cell(ws.cell(row=r, column=col))

    # Auto-fit column widths based on header + data content
    for i, h in enumerate(HEADERS, start=1):
        values = [str(r.get(h, "")) for r in agg_rows]
        width  = max(len(h), *(len(v) for v in values)) if values else len(h)
        ws.column_dimensions[get_column_letter(i)].width = min(max(width + 3, 13), 32)

    _apply_autofilter(ws, n_cols)
    ws.freeze_panes = "A3"


def fetch_flex():
    """Download the Flex statement. Returns (trade_rows, flex_account, error).
    A failure is reported, not fatal: the report is still built from the store."""
    try:
        ref      = send_request()
        xml_text = get_statement(ref)
        return parse_trades(xml_text), parse_account_id(xml_text), None
    except (requests.RequestException, RuntimeError, ET.ParseError) as e:
        print(f"[Flex] download failed: {e} - building the report from stored "
              f"fills only.", file=sys.stderr)
        return [], None, str(e)


def resolve_account(managed_accounts, flex_account):
    """IBKR_ACCOUNT from .env, else the Gateway login's single account, else the
    Flex statement's account."""
    if ACCOUNT:
        return ACCOUNT
    if len(managed_accounts) == 1:
        return managed_accounts[0]
    if len(managed_accounts) > 1:
        sys.exit(f"Error: the Gateway login has several accounts {managed_accounts}. "
                 f"Set IBKR_ACCOUNT in .env to the one to report on.")
    if flex_account and flex_account != "N/A":
        return flex_account
    sys.exit("Error: could not determine the account (Gateway not reachable and no "
             "Flex data). Set IBKR_ACCOUNT in .env.")


def _for_account(acct):
    return lambda a: (a or "") == acct


def main():
    warnings = []

    print("Connecting to TWS - fetching open/completed orders, executions, and account summary...")
    (orders_by_symbol, account_data, executions,
     completed_orders, managed_accounts, portfolio) = fetch_tws_data()
    if not managed_accounts and not executions and not orders_by_symbol:
        warnings.append("gateway-unreachable-or-empty")

    print("Downloading Flex statement...")
    flex_rows, flex_account, flex_error = fetch_flex()
    if flex_error:
        warnings.append("flex-failed")

    set_account(resolve_account(managed_accounts, flex_account))
    print(f"Reporting account: {ACCOUNT}")

    # ── Keep only the reported account's data from both feeds ─────────────
    mine = _for_account(ACCOUNT)
    if managed_accounts and ACCOUNT not in managed_accounts:
        print(f"[WARN] Gateway is logged in to {managed_accounts}, not {ACCOUNT} - "
              f"no live data for this account.", file=sys.stderr)
        warnings.append("gateway-other-account")
    if flex_account and flex_account not in ("N/A", ACCOUNT):
        print(f"[WARN] The Flex Query belongs to {flex_account}, not {ACCOUNT} - its "
              f"trades are ignored. Create the Flex Query under {ACCOUNT}.", file=sys.stderr)
        warnings.append("flex-other-account")

    executions = {k: v for k, v in executions.items() if mine(v.get("account"))}
    orders_by_symbol = {sym: legs for sym, legs in
                        ((s, [o for o in l if not o.get("account") or mine(o["account"])])
                         for s, l in orders_by_symbol.items()) if legs}
    completed_orders = [o for o in completed_orders
                        if not o.get("account") or mine(o["account"])]
    account_data = account_data.get(ACCOUNT, {})
    if portfolio is not None:
        portfolio = {k: v for k, v in portfolio.items() if mine(v.get("account"))}
        print(f"  Live: {len(portfolio)} open position(s) with market prices.")
    else:
        print("[WARN] No live positions from Gateway - Unrealized P&L will be blank.",
              file=sys.stderr)
        warnings.append("no-portfolio")
    flex_mine = [r for r in flex_rows if mine(r.get("accountId") or flex_account)]

    link_parent_orders(orders_by_symbol)
    pending_rows = build_pending_rows(orders_by_symbol)
    print(f"  Live: {len(pending_rows)} pending order rows | {len(executions)} fill(s) today | "
          f"{len(completed_orders)} completed order(s) | {len(account_data)} account tag(s).")
    print(f"  Flex: {len(flex_mine)} of {len(flex_rows)} trade record(s) belong to {ACCOUNT}.")

    # ── Report start date ─────────────────────────────────────────────────
    state = load_state()
    if _DATA_FROM_ENV:
        try:
            start = dt.date.fromisoformat(_DATA_FROM_ENV)
        except ValueError:
            sys.exit(f"Error: IBKR_DATA_FROM must be YYYY-MM-DD, got {_DATA_FROM_ENV!r}.")
    else:
        start = dt.date.fromisoformat(state.setdefault("start_date", dt.date.today().isoformat()))
    set_data_from(start)
    print(f"  Report covers trades from {DATA_FROM:%d-%b-%Y}.")

    # ── Order ledger (trigger / limit per order) ──────────────────────────
    ledger = load_order_ledger()
    store  = load_fill_store()
    if not state.get("legacy_imported"):
        import_legacy_stores(store, ledger)
        state["legacy_imported"] = dt.date.today().isoformat()
    update_order_ledger(ledger, orders_by_symbol, completed_orders)
    save_order_ledger(ledger)
    order_lookup = build_order_lookup(ledger)
    print(f"  [Order Ledger] {len(ledger)} order(s) tracked ({_ledger_file()}).")

    # ── Fill store: live first, then Flex (Flex enriches/confirms) ────────
    counts = defaultdict(int)
    for ex in executions.values():
        counts["live_" + upsert_fill(store, _exec_to_flex_row(ex), "LIVE")] += 1
    for r in flex_mine:
        counts["flex_" + upsert_fill(store, r, "FLEX")] += 1

    # Legacy fills with no execution ID can't be matched to Flex by ID. Live
    # fills always carry an ID now, so these only come from the old JSON stores.
    # Inside the period the Flex statement covers, Flex is the complete record,
    # so such leftovers would be double counts - drop them.
    flex_dates = [d for d in (parse_trade_date(r) for r in flex_mine) if d]
    if flex_dates:
        lo, hi = min(flex_dates), max(flex_dates)
        stale = [k for k, r in store.items()
                 if k.startswith("fp:") and r.get("source") == "LIVE"
                 and lo <= (parse_trade_date(r) or dt.date.min) <= hi]
        for k in stale:
            del store[k]
        if stale:
            print(f"  [Cleanup] removed {len(stale)} old no-ID fill(s) already covered by Flex.")
    days_written = save_fill_store(store)
    state["last_run"] = dt.datetime.now().isoformat(timespec="seconds")
    save_state(state)
    print(f"  [Fill Store] {len(store)} fill(s) stored | live new={counts['live_new']} "
          f"merged={counts['live_merged']} | flex new={counts['flex_new']} "
          f"merged={counts['flex_merged']} | {days_written} day file(s) updated "
          f"({_fills_dir()}).")

    unconfirmed = [r for r in store.values() if r.get("source") == "LIVE"
                   and (parse_trade_date(r) or dt.date.max) < dt.date.today() - dt.timedelta(days=3)]
    if unconfirmed and not flex_error:
        print(f"[WARN] {len(unconfirmed)} live fill(s) older than 3 days are still not in "
              f"the Flex statement - check the Flex Query's account and period.",
              file=sys.stderr)

    rows = list(store.values())
    if not rows:
        print("No trades stored yet for this account.")
    write_excel(rows, pending_rows, [], ACCOUNT, account_data, order_lookup,
                portfolio=portfolio)

    log_run("WARN" if warnings else "OK",
            fills_stored=len(store), live_today=len(executions), flex_rows=len(flex_mine),
            pending=len(pending_rows), report=os.path.basename(OUTPUT_FILE),
            warnings=",".join(warnings) or "-")


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        if e.code not in (None, 0):
            log_run("FAILED", error=str(e.code).splitlines()[0])
        raise
    except Exception as e:                                  # noqa: BLE001
        log_run("FAILED", error=f"{type(e).__name__}: {e}")
        raise
