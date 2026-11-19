"""
Standalone diagnostic: connect to TWS and print ONLY the live open positions.
Isolates the reqPositions feed from the rest of the report so we can see exactly
what IBKR returns.

    python test_positions.py
"""
import os
import threading
import time

from ibapi.client import EClient
from ibapi.wrapper import EWrapper

# This script is run directly (not via run.ps1), so load .env ourselves if it's
# present, so IBKR_TWS_PORT etc. are picked up the same way the report gets them.
def _load_dotenv(path=".env"):
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            os.environ.setdefault(key.strip(), val.strip())

_load_dotenv()

# Same env vars as the report (set IBKR_TWS_PORT=7497 for paper). CLIENT_ID is
# kept distinct from the report's so the two can connect at the same time.
HOST      = os.environ.get("IBKR_TWS_HOST", "127.0.0.1")
PORT      = int(os.environ.get("IBKR_TWS_PORT", "7496"))  # 7496 live TWS | 7497 paper | 4001/4002 Gateway
CLIENT_ID = 997         # different from the report's 998 to avoid a clash


class PosApp(EWrapper, EClient):
    def __init__(self):
        EClient.__init__(self, self)
        self.positions = []
        self._done = threading.Event()
        self.connected_ok = False

    def error(self, reqId, code, msg, advancedOrderRejectJson=""):
        # Print everything so we don't hide a relevant warning/rejection.
        print(f"[TWS] id={reqId} code={code}: {msg}")

    def connectAck(self):
        self.connected_ok = True
        print("[TWS] connectAck - socket connected.")

    def nextValidId(self, orderId):
        self.connected_ok = True
        print(f"[TWS] nextValidId={orderId} - requesting positions...")
        self.reqPositions()

    def position(self, account, contract, position, avgCost):
        print(f"  POSITION  acct={account}  symbol={contract.symbol}  "
              f"secType={contract.secType}  ccy={contract.currency}  "
              f"exch={contract.exchange or contract.primaryExchange}  "
              f"localSymbol={contract.localSymbol}  qty={position}  avgCost={avgCost}")
        self.positions.append((account, contract.symbol, position, avgCost))

    def positionEnd(self):
        print(f"[TWS] positionEnd - {len(self.positions)} position(s) total.")
        self._done.set()


def main():
    app = PosApp()
    print(f"Connecting to {HOST}:{PORT} (clientId={CLIENT_ID})...")
    try:
        app.connect(HOST, PORT, CLIENT_ID)
    except Exception as e:
        print(f"[TWS] connect failed: {e}")
        return
    threading.Thread(target=app.run, daemon=True).start()

    got = app._done.wait(timeout=15)
    if not got:
        print("[TWS] TIMEOUT - positionEnd never arrived within 15s.")
    print(f"\nconnected_ok={app.connected_ok}  positions_received={len(app.positions)}")
    if app.positions:
        nonzero = [p for p in app.positions if p[2]]
        print(f"non-zero (open) positions: {len(nonzero)}")
        for acct, sym, qty, cost in nonzero:
            print(f"   OPEN  {sym:<10} qty={qty}  avgCost={cost}  acct={acct}")
    else:
        print("No positions returned by TWS.")

    time.sleep(0.5)
    app.disconnect()


if __name__ == "__main__":
    main()
