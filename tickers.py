"""
tickers.py — Shared PSX portfolio universe used by both automations
(monthly price report and dividend tracker).

Symbols verified live against https://dps.psx.com.pk/company/<symbol> on
2026-07-10. If PSX renames/delists a symbol, update this list — nothing
else needs to change since both automations import from here.

NB: several source symbols were company names rather than tickers and were
resolved to their real PSX symbol (e.g. "Bank Al falah" -> BAFL, "Kohat
Cement" -> KOHC, "Agha Steel" -> AGHA, "LUCKY" -> LUCK). "ENGRO holding"
maps to ENGROH (Engro Holdings Limited) — Engro Corporation Limited (the
old ENGRO symbol) stopped trading in Jan 2025 following a group
restructuring; ENGROH is the live successor.

The "Additional Holdings" tickers below were added because they appear as
real positions in Cash Dividend 1 (1).xlsx (see holdings.py) but weren't
part of the original 9-sector universe. "PANTHER" in that file resolves
to PTL (Panther Tyres Ltd.) and "PACKAGE" resolves to PKGS (Packages
Limited) — both symbols verified live against PSX.
"""

UNIVERSE = [
    # Banks
    {"ticker": "MCB", "company": "MCB Bank Limited", "sector": "Banks"},
    {"ticker": "HBL", "company": "Habib Bank Limited", "sector": "Banks"},
    {"ticker": "BAFL", "company": "Bank Alfalah Limited", "sector": "Banks"},
    {"ticker": "BAHL", "company": "Bank AL Habib Limited", "sector": "Banks"},
    {"ticker": "MEBL", "company": "Meezan Bank Limited", "sector": "Banks"},
    {"ticker": "HMB", "company": "Habib Metropolitan Bank Limited", "sector": "Banks"},
    {"ticker": "UBL", "company": "United Bank Limited", "sector": "Banks"},
    # Cement
    {"ticker": "LUCK", "company": "Lucky Cement Limited", "sector": "Cement"},
    {"ticker": "KOHC", "company": "Kohat Cement Company Limited", "sector": "Cement"},
    {"ticker": "DGKC", "company": "D.G. Khan Cement Company Limited", "sector": "Cement"},
    {"ticker": "CHCC", "company": "Cherat Cement Company Limited", "sector": "Cement"},
    {"ticker": "FCCL", "company": "Fauji Cement Company Limited", "sector": "Cement"},
    {"ticker": "MLCF", "company": "Maple Leaf Cement Factory Limited", "sector": "Cement"},
    # Fertilizers
    {"ticker": "ENGROH", "company": "Engro Holdings Limited", "sector": "Fertilizers"},
    {"ticker": "EFERT", "company": "Engro Fertilizers Limited", "sector": "Fertilizers"},
    {"ticker": "FATIMA", "company": "Fatima Fertilizer Company Limited", "sector": "Fertilizers"},
    {"ticker": "FFC", "company": "Fauji Fertilizer Company Limited", "sector": "Fertilizers"},
    # Oil/Gas
    {"ticker": "POL", "company": "Pakistan Oilfields Limited", "sector": "Oil/Gas"},
    {"ticker": "PSO", "company": "Pakistan State Oil Company Limited", "sector": "Oil/Gas"},
    {"ticker": "MARI", "company": "Mari Energies Limited", "sector": "Oil/Gas"},
    {"ticker": "PPL", "company": "Pakistan Petroleum Limited", "sector": "Oil/Gas"},
    {"ticker": "OGDC", "company": "Oil & Gas Development Company Limited", "sector": "Oil/Gas"},
    # Power
    {"ticker": "ALTN", "company": "Altern Energy Limited", "sector": "Power"},
    {"ticker": "NCPL", "company": "Nishat Chunian Power Limited", "sector": "Power"},
    {"ticker": "KOHE", "company": "Kohinoor Energy Limited", "sector": "Power"},
    {"ticker": "NPL", "company": "Nishat Power Limited", "sector": "Power"},
    {"ticker": "HUBC", "company": "The Hub Power Company Limited", "sector": "Power"},
    # Steel Sector
    {"ticker": "ISL", "company": "International Steels Limited", "sector": "Steel Sector"},
    {"ticker": "AGHA", "company": "Agha Steel Industries Limited", "sector": "Steel Sector"},
    {"ticker": "ASTL", "company": "Amreli Steels Limited", "sector": "Steel Sector"},
    {"ticker": "ASL", "company": "Aisha Steel Mills Limited", "sector": "Steel Sector"},
    # Others
    {"ticker": "CEPB", "company": "Century Paper & Board Mills Limited", "sector": "Others"},
    {"ticker": "INDU", "company": "Indus Motor Company Limited", "sector": "Others"},
    {"ticker": "NML", "company": "Nishat Mills Limited", "sector": "Others"},
    {"ticker": "AGIL", "company": "Agriauto Industries Limited", "sector": "Others"},
    {"ticker": "OLPL", "company": "OLP Financial Services Pakistan Limited", "sector": "Others"},
    {"ticker": "EPCL", "company": "Engro Polymer & Chemicals Limited", "sector": "Others"},
    # Pharma
    {"ticker": "ABOT", "company": "Abbott Laboratories (Pakistan) Limited", "sector": "Pharma"},
    {"ticker": "FEROZ", "company": "Ferozsons Laboratories Limited", "sector": "Pharma"},
    {"ticker": "GLAXO", "company": "GlaxoSmithKline Pakistan Limited", "sector": "Pharma"},
    {"ticker": "HINOON", "company": "Highnoon Laboratories Limited", "sector": "Pharma"},
    {"ticker": "CPHL", "company": "Citi Pharma Limited", "sector": "Pharma"},
    {"ticker": "AGP", "company": "AGP Limited", "sector": "Pharma"},
    # Tech
    {"ticker": "SYS", "company": "Systems Limited", "sector": "Tech"},
    {"ticker": "TRG", "company": "TRG Pakistan Limited", "sector": "Tech"},
    {"ticker": "AVN", "company": "Avanceon Limited", "sector": "Tech"},
    # Additional Holdings (from Cash Dividend 1 (1).xlsx, not in the original 9 sectors)
    {"ticker": "AICL", "company": "Adamjee Insurance Company Limited", "sector": "Insurance"},
    {"ticker": "PTL", "company": "Panther Tyres Ltd.", "sector": "Automobile Parts & Accessories"},
    {"ticker": "BYCO", "company": "Byco Petroleum Pakistan Limited", "sector": "Refinery"},
    {"ticker": "CPPL", "company": "Cherat Packaging Limited", "sector": "Paper & Board"},
    {"ticker": "FDIBL", "company": "First Dawood Investment Bank Limited", "sector": "Investment Banks / Companies"},
    {"ticker": "SRVI", "company": "Service Industries Limited", "sector": "Leather & Tanneries"},
    {"ticker": "PKGS", "company": "Packages Limited", "sector": "Paper & Board"},
    {"ticker": "AKBL", "company": "Askari Bank Limited", "sector": "Banks"},
    {"ticker": "MUREB", "company": "Murree Brewery Company Limited", "sector": "Food & Personal Care Products"},
    {"ticker": "GGL", "company": "Ghani Global Holdings Limited", "sector": "Miscellaneous"},
    {"ticker": "TPL", "company": "TPL Corp Limited", "sector": "Technology & Communication"},
    {"ticker": "BFAGRO", "company": "Barkat Frisian Agro Limited", "sector": "Food & Personal Care Products"},
    {"ticker": "ATLH", "company": "Atlas Honda Limited", "sector": "Automobile Assembler"},
    {"ticker": "PIOC", "company": "Pioneer Cement Limited", "sector": "Cement"},
    {"ticker": "SRR", "company": "Signature Residency Reit", "sector": "Real Estate Investment Trust"},
]

TICKERS = [row["ticker"] for row in UNIVERSE]
COMPANY_NAMES = {row["ticker"]: row["company"] for row in UNIVERSE}
SECTORS = {row["ticker"]: row["sector"] for row in UNIVERSE}

# Face value (PKR) per share, used to convert PSX payout percentages into
# per-share rupee amounts. Nearly all PSX equities have a Rs. 10 face
# value; override here if a specific ticker differs.
FACE_VALUE = {ticker: 10.0 for ticker in TICKERS}

# Automation 2 (dividend tracker) uses a narrower universe than the price
# report: only tickers actually held by Ellahi Capital/AATML, confirmed by
# the user. This intentionally excludes tickers that are in the 9-sector
# price-tracking universe but aren't held (e.g. KOHC, POL, ISL, ABOT) —
# the dividend tracker must not report on those. Source-list aliases:
# PANTHER -> PTL, "Aisha Steel"/"Aissha Steel" -> ASL, PACKAGE -> PKGS,
# "HINOON LAB" -> HINOON, MAPLE -> MLCF, LUCKY/Luck -> LUCK (same
# resolutions as tickers.py's own docstring / holdings.py's ALIASES).
DIVIDEND_TICKERS = [
    "UBL", "HBL", "BAHL", "BAFL", "HMB", "MCB",
    "ALTN", "AGIL", "AICL", "FFC", "FCCL", "PTL", "ASL", "BYCO", "CPPL",
    "FDIBL", "SRVI", "NPL", "PKGS", "AKBL", "MEBL", "HINOON", "MUREB",
    "GGL", "MLCF", "TPL", "AGP", "BFAGRO", "DGKC", "SYS", "LUCK", "ATLH",
    "PIOC", "SRR", "OGDC", "PPL", "NCPL",
]
DIVIDEND_UNIVERSE = [row for row in UNIVERSE if row["ticker"] in DIVIDEND_TICKERS]
