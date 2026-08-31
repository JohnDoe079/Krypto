# config.py
"""Konfiguracja nazw arkuszy i mapowania kolumn."""

BINANCE_SHEETS = {
    "Customer Information": "customer_info",
    "KYC Documents": "kyc_docs",
    "Assets Overview": "assets_overview",
    "Spot Asset Log": "spot_asset_log",
    "Funding Asset Log": "funding_asset_log",
    "Fiat Deposit History": "fiat_deposit",
    "Fiat Trades": "fiat_trades",
    "Deposit History": "deposit_history",
    "Withdrawal History": "withdrawal_history",
    "Attempted Withdrawal History": "attempted_withdrawal",
    "Binance Pay": "binance_pay",
    "P2P": "p2p",
    "OTC Trading": "otc_trading",
    "Access Logs": "access_logs",
    "Order History": "order_history",
    "Approved Devices": "approved_devices",
}

HTX_SHEETS = {
    "register_1": "register_1",
    "balance_1": "balance_1",
    "login_1": "login_1",
}

DATA_DIRS = {
    "binance": "data/binance",
    "htx": "data/htx",
}
