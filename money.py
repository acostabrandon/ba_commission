"""Currency rounding that matches Excel (round half up to the cent)."""
from decimal import Decimal, ROUND_HALF_UP


def r2(x) -> float:
    if x is None:
        return None
    return float(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
