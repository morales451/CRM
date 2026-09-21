"""Warranty-rate calculator — application rates and material quantities.

Ported verbatim from the Warranty Roofing Calculator
(warrantyroofingcalculator.netlify.app) so the roof report quotes the same
numbers the calculator does.

Rates are gallons per roofing square (1 square = 100 sq ft) and vary by
coating system, roof surface, and the warranty length being quoted.
"""

import math

# system -> roof type -> warranty years -> gallons per square, per coat
RATES = {
    "Silicone": {
        "Capsheet":    {10: {"base": 1.25, "top1": 2,   "top2": 0,    "top3": 0},
                        15: {"base": 1.25, "top1": 2.5, "top2": 0,    "top3": 0},
                        20: {"base": 1.25, "top1": 3,   "top2": 0,    "top3": 0}},
        "Sprayfoam":   {10: {"base": 0, "top1": 1.5, "top2": 0, "top3": 0},
                        15: {"base": 0, "top1": 2,   "top2": 0, "top3": 0},
                        20: {"base": 0, "top1": 2.5, "top2": 0, "top3": 0}},
        "Single-Ply":  {10: {"base": 0, "top1": 1.5, "top2": 0, "top3": 0},
                        15: {"base": 0, "top1": 2,   "top2": 0, "top3": 0},
                        20: {"base": 0, "top1": 2.5, "top2": 0, "top3": 0}},
        "Metal":       {10: {"base": 0, "top1": 1.5, "top2": 0, "top3": 0},
                        15: {"base": 0, "top1": 2,   "top2": 0, "top3": 0},
                        20: {"base": 0, "top1": 2.5, "top2": 0, "top3": 0}},
    },
    "Acrylic": {
        "Reinforced": {
            "Capsheet":   {10: {"base": 2, "top1": 2, "top2": 1.75, "top3": 0},
                           15: {"base": 2, "top1": 2, "top2": 1.5,  "top3": 1.5},
                           20: {"base": 2, "top1": 2, "top2": 2,    "top3": 2}},
            "Single-Ply": {10: {"base": 2, "top1": 1, "top2": 1.5,  "top3": 0},
                           15: {"base": 2, "top1": 1, "top2": 1.5,  "top3": 1.5},
                           20: {"base": 2, "top1": 1, "top2": 2,    "top3": 2}},
            "Sprayfoam": None,
            "Metal": None,
        },
        "Standard": {
            "Capsheet":   {10: {"base": 1.75, "top1": 2,   "top2": 0,   "top3": 0},
                           15: {"base": 2,    "top1": 2,   "top2": 0,   "top3": 0},
                           20: {"base": 2,    "top1": 2,   "top2": 1.5, "top3": 0}},
            "Sprayfoam":  {10: {"base": 1.5, "top1": 1.5, "top2": 0, "top3": 0},
                           15: {"base": 1.5, "top1": 2,   "top2": 0, "top3": 0},
                           20: {"base": 1.5, "top1": 1.5, "top2": 2, "top3": 0}},
            "Single-Ply": {10: {"base": 1.5, "top1": 1.5, "top2": 0, "top3": 0},
                           15: {"base": 1.5, "top1": 2,   "top2": 0, "top3": 0},
                           20: {"base": 1.5, "top1": 1.5, "top2": 2, "top3": 0}},
            "Metal":      {10: {"base": 1.5, "top1": 1.5, "top2": 0, "top3": 0},
                           15: {"base": 1.5, "top1": 2,   "top2": 0, "top3": 0},
                           20: {"base": 1.5, "top1": 1.5, "top2": 2, "top3": 0}},
        },
    },
    "Aluminum": {
        "Metal":      {10: {"base": 0, "top1": 2,   "top2": 0, "top3": 0}, 15: None, 20: None},
        "Capsheet":   {10: {"base": 0, "top1": 2.5, "top2": 0, "top3": 0}, 15: None, 20: None},
        "Sprayfoam": None,
        "Single-Ply": None,
    },
}

COATING_SYSTEMS = ["Silicone", "Acrylic", "Aluminum"]
ACRYLIC_TYPES = ["Standard", "Reinforced"]
ROOF_TYPES = ["Capsheet", "Single-Ply", "Sprayfoam", "Metal"]
WARRANTY_YEARS = [10, 15, 20]

ADHESION_PRIMER_RATE = 0.2   # gal/square when the adhesion test fails
RUST_PRIMER_RATE = 0.5       # gal/square for a full rust prime on metal
FASTENER_CAULK_PER_TUBE = 125
PAIL_GALLONS = 5

# Product catalog, ported from the calculator. Henry Prograde names are the
# defaults (listed first); Enduraroof equivalents are the alternates.
PRODUCT_CATALOG = {
    "Silicone": {
        "topcoats": ["Prograde 988 Silicone", "Enduraroof Premium Silicone"],
        "basecoats": ["Prograde 294 BaseCoat", "Enduraroof BaseCoat & Sealer"],
        "butter_grades": ["Prograde 923 Butter Grade", "EnduraRoof Butter Grade"],
        "fabrics": ["Prograde 195 (SOFT)", "Prograde 196 (FIRM)",
                    "Enduraroof Polyester Fabric"],
    },
    "Acrylic": {
        "topcoats": ["Acryshield 400", "Acryshield 510", "Acryshield 550HT",
                     "Acryshield 610", "Enduraroof Elastomeric",
                     "Enduraroof Premium Acrylic"],
        "basecoats": ["Acryshield Basecoat", "Acryshield 400", "Acryshield 505",
                      "Enduraroof Basecoat", "Enduraroof Elastomeric",
                      "Enduraroof Premium Acrylic"],
        "butter_grades": ["Prograde 289 White Roofing Sealant",
                          "Prograde 295 Metal Seam Sealer",
                          "Enduraroof Acrylic Roof Patch"],
        "fabrics": ["Prograde 195 (SOFT)", "Prograde 196 (FIRM)",
                    "Enduraroof Polyester Fabric"],
    },
    "Aluminum": {
        "topcoats": ["Pro-Grade 586", "Enduraroof Fibered Aluminum"],
        "basecoats": [],
        "butter_grades": ["Prograde 289 White Roofing Sealant",
                          "Enduraroof Acrylic Roof Patch"],
        "fabrics": ["Prograde 195 (SOFT)", "Prograde 196 (FIRM)",
                    "Enduraroof Polyester Fabric"],
    },
}

# Primer names and the warranty holder follow the brand of the chosen coating.
PRIMERS = {
    "Prograde": {"adhesion": "Prograde 941 Adhesion Promoting Primer",
                 "rust": "PrimeTek Rust Inhibiting Primer"},
    "Enduraroof": {"adhesion": "Enduraroof Silicone Roof Primer",
                   "rust": "Enduraroof Metal Roofing Primer"},
}

MANUFACTURERS = {
    "Prograde": "Henry Company (a Carlisle Company)",
    "Enduraroof": "Enduraroof",
}

# Short form for running prose ("a 10-year warranty from Henry")
MANUFACTURERS_SHORT = {"Prograde": "Henry", "Enduraroof": "Enduraroof"}


def brand_of(product: str) -> str:
    """Which product line a name belongs to (the calculator's own rule)."""
    return "Enduraroof" if product and "Endura" in product else "Prograde"


def default_products(coating_system: str) -> dict:
    cat = PRODUCT_CATALOG.get(coating_system, PRODUCT_CATALOG["Silicone"])
    return {
        "top": cat["topcoats"][0] if cat["topcoats"] else "",
        "base": cat["basecoats"][0] if cat["basecoats"] else "",
        "mastic": cat["butter_grades"][0] if cat["butter_grades"] else "",
    }


def resolve_products(coating_system: str, topcoat="", basecoat="", butter_grade=""):
    """Chosen products, falling back to the system's defaults. Anything not in
    the catalog for this system is replaced, so switching systems can't leave
    a silicone product on an aluminum job."""
    cat = PRODUCT_CATALOG.get(coating_system, PRODUCT_CATALOG["Silicone"])
    defaults = default_products(coating_system)
    top = topcoat if topcoat in cat["topcoats"] else defaults["top"]
    base = basecoat if basecoat in cat["basecoats"] else defaults["base"]
    mastic = butter_grade if butter_grade in cat["butter_grades"] else defaults["mastic"]
    brand = brand_of(top)
    return {"top": top, "base": base, "mastic": mastic, "brand": brand,
            "adhesion_primer": PRIMERS[brand]["adhesion"],
            "rust_primer": PRIMERS[brand]["rust"],
            "manufacturer": MANUFACTURERS[brand],
            "manufacturer_short": MANUFACTURERS_SHORT[brand]}


def supported_roof_types(coating_system: str, acrylic_system_type: str = "Standard"):
    """Roof types this system can actually be installed over."""
    if coating_system == "Acrylic":
        table = RATES["Acrylic"].get(acrylic_system_type or "Standard", {})
    else:
        table = RATES.get(coating_system, {})
    return [rt for rt in ROOF_TYPES if table.get(rt)]


def supported_warranties(coating_system: str, roof_type: str,
                         acrylic_system_type: str = "Standard"):
    """Warranty lengths published for this system on this roof."""
    return [y for y in WARRANTY_YEARS
            if get_rates(coating_system, roof_type, y, acrylic_system_type)]


def round_to_pails(gallons: float) -> int:
    """Round up to whole 5-gallon pails (the calculator's `at` helper)."""
    return 0 if gallons <= 0 else math.ceil(gallons / PAIL_GALLONS) * PAIL_GALLONS


def get_rates(coating_system: str, roof_type: str, warranty_years: int,
              acrylic_system_type: str = "Standard"):
    """Gallons-per-square for each coat, or None if the system can't carry
    that warranty on that roof (e.g. Aluminum over 10 years)."""
    if coating_system == "Acrylic":
        table = (RATES["Acrylic"].get(acrylic_system_type or "Standard") or {}).get(roof_type)
    else:
        table = RATES.get(coating_system, {}).get(roof_type)
    if not table:
        return None
    return table.get(int(warranty_years))


def calculate(sqft, coating_system="Silicone", roof_type="Capsheet",
              warranty_years=10, acrylic_system_type="Standard",
              deduction_sqft=0, linear_feet=0, waste_pct=0, stretch_pct=0,
              passed_adhesion=True, has_rust=False, rust_prime_method="field",
              topcoat="", basecoat="", butter_grade=""):
    """Materials plan for one roof. Percentages are whole numbers (5 = 5%).

    Returns None when the chosen system/roof/warranty combination has no
    published rate; otherwise a dict with per-coat rates, gallons and pails,
    primers, accessories, and totals.
    """
    net_sqft = max(0, (sqft or 0) - (deduction_sqft or 0))
    squares = net_sqft / 100 if net_sqft > 0 else 0
    factor = 1 + (waste_pct or 0) / 100 + (stretch_pct or 0) / 100

    rates = get_rates(coating_system, roof_type, warranty_years, acrylic_system_type)
    if not rates:
        return None

    products = resolve_products(coating_system, topcoat, basecoat, butter_grade)
    coats = []
    for key, label, short in (("base", "Basecoat", "Base"),
                              ("top1", "Topcoat", "Top"),
                              ("top2", "Second Topcoat", "Top 2"),
                              ("top3", "Third Topcoat", "Top 3")):
        rate = rates.get(key) or 0
        if not rate:
            continue
        gallons = round_to_pails(squares * rate * factor)
        coats.append({
            "key": key,
            "label": label,
            "short": short,
            "product": products["base"] if key == "base" else products["top"],
            "rate": rate,
            "raw_gallons": round(squares * rate * factor, 1),
            "gallons": gallons,
            "pails": gallons // PAIL_GALLONS,
        })

    adhesion_primer = (0 if passed_adhesion
                       else math.ceil(squares * ADHESION_PRIMER_RATE * factor))
    rust_primer = 0
    if (coating_system in ("Silicone", "Acrylic") and roof_type == "Metal"
            and has_rust and rust_prime_method != "spot"):
        rust_primer = round_to_pails(squares * RUST_PRIMER_RATE * factor)

    # Butter-grade mastic for seams and penetrations, by linear feet
    lf_per_container = 150 if coating_system == "Acrylic" else 80
    mastic_buckets = math.ceil(linear_feet / lf_per_container) if linear_feet else 0
    fastener_caulk_tubes = 0
    estimated_screws = 0
    if roof_type == "Metal" and net_sqft > 0:
        estimated_screws = math.ceil(net_sqft * 0.8)
        fastener_caulk_tubes = math.ceil(estimated_screws / FASTENER_CAULK_PER_TUBE)

    membrane_rolls = 0
    if coating_system == "Acrylic" and acrylic_system_type == "Reinforced":
        membrane_rolls = math.ceil(net_sqft * factor / 1080) if net_sqft else 0

    total_gallons = sum(c["gallons"] for c in coats) + adhesion_primer + rust_primer
    return {
        "coating_system": coating_system,
        "acrylic_system_type": acrylic_system_type,
        "roof_type": roof_type,
        "warranty_years": int(warranty_years),
        "net_sqft": net_sqft,
        "squares": round(squares, 1),
        "adjusted_squares": round(squares * factor, 1),
        "factor": round(factor, 4),
        "waste_pct": waste_pct or 0,
        "stretch_pct": stretch_pct or 0,
        "coats": coats,
        "adhesion_primer_gal": adhesion_primer,
        "rust_primer_gal": rust_primer,
        "mastic_buckets": mastic_buckets,
        "mastic_product": products["mastic"],
        "mastic_container": "3.5-gal" if coating_system == "Acrylic" else "2-gal",
        "lf_per_container": lf_per_container,
        "membrane_rolls": membrane_rolls,
        "estimated_screws": estimated_screws,
        "fastener_caulk_tubes": fastener_caulk_tubes,
        "total_gallons": total_gallons,
        "total_pails": math.ceil(total_gallons / PAIL_GALLONS) if total_gallons else 0,
        "products": products,
    }


def warranty_options(sqft, coating_system, roof_type, **kwargs):
    """The same roof priced at 10/15/20 years — the calculator's comparison."""
    out = []
    for years in WARRANTY_YEARS:
        plan = calculate(sqft, coating_system=coating_system, roof_type=roof_type,
                         warranty_years=years, **kwargs)
        if plan:
            out.append(plan)
    return out


def guess_roof_type(surface_type: str) -> str:
    """Map a free-text surface description to a calculator roof type."""
    s = (surface_type or "").lower()
    if "foam" in s or "spf" in s:
        return "Sprayfoam"
    if "metal" in s or "standing seam" in s or "r-panel" in s:
        return "Metal"
    if any(k in s for k in ("tpo", "epdm", "pvc", "single", "membrane", "ply")):
        return "Single-Ply"
    return "Capsheet"


# ---------------------------------------------------------------- Pricing

# Sell price per square foot of coated roof. Longer warranties need more
# coating, so each step up adds to the base rate (cumulative at 20 years).
DEFAULT_PRICING = {
    "capsheet_base": 4.50,   # capsheet, 10-year system
    "other_base": 4.00,      # single-ply, sprayfoam, metal — 10-year system
    "add_15": 0.15,          # added for a 15-year system
    "add_20": 0.10,          # added on top of add_15 for a 20-year system
}


def price_per_sqft(roof_type: str, warranty_years, pricing: dict | None = None) -> float:
    p = {**DEFAULT_PRICING, **(pricing or {})}
    rate = p["capsheet_base"] if roof_type == "Capsheet" else p["other_base"]
    years = int(warranty_years or 10)
    if years >= 15:
        rate += p["add_15"]
    if years >= 20:
        rate += p["add_20"]
    return round(rate, 2)


def suggested_price(net_sqft, roof_type: str, warranty_years,
                    pricing: dict | None = None) -> dict:
    """Suggested sell price for a roof, priced on the coated (net) area."""
    rate = price_per_sqft(roof_type, warranty_years, pricing)
    sqft = max(0, net_sqft or 0)
    return {"rate": rate, "sqft": sqft, "total": round(sqft * rate, 2),
            "warranty_years": int(warranty_years or 10), "roof_type": roof_type}
