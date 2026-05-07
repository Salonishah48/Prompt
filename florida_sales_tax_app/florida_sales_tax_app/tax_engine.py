"""
Florida DR-15 Sales & Use Tax Calculation Engine
=================================================
Implements the rules from DR-15N (R. 10/25) and the Florida Sales Tax
Compliance Guide. All rules are loaded from rules/tax_rules.json so they
can be updated without changing code.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from decimal import Decimal, ROUND_HALF_UP, ROUND_UP
from pathlib import Path
from typing import Optional


# ---------- utility ----------
def money(value) -> Decimal:
    """Convert to Decimal and quantize to cents using Florida's rounding rule:
    compute to 3 decimal places; if the third decimal is > 4, round UP to cent."""
    d = Decimal(str(value)) if not isinstance(value, Decimal) else value
    # carry to 3 decimals
    three = d.quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)
    third = int((three * 1000) % 10)
    if third > 4:
        return three.quantize(Decimal("0.01"), rounding=ROUND_UP)
    return three.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def simple_round(value) -> Decimal:
    d = Decimal(str(value)) if not isinstance(value, Decimal) else value
    return d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


# ---------- data classes for the return ----------
@dataclass
class LineResult:
    """One row of the DR-15 front panel (Lines A–E)."""
    label: str
    gross_sales: Decimal = Decimal("0.00")
    exempt_sales: Decimal = Decimal("0.00")
    taxable_amount: Decimal = Decimal("0.00")
    tax_due: Decimal = Decimal("0.00")          # includes surtax
    state_tax: Decimal = Decimal("0.00")        # state portion only
    surtax: Decimal = Decimal("0.00")           # county portion only


@dataclass
class DR15Return:
    reporting_period: str = ""
    county: str = ""
    surtax_rate: Decimal = Decimal("0")

    line_a: LineResult = field(default_factory=lambda: LineResult("A. Sales/Services/Electricity"))
    line_b: LineResult = field(default_factory=lambda: LineResult("B. Taxable Purchases (Use Tax)"))
    line_c: LineResult = field(default_factory=lambda: LineResult("C. Commercial Rentals"))
    line_d: LineResult = field(default_factory=lambda: LineResult("D. Transient Rentals"))
    line_e: LineResult = field(default_factory=lambda: LineResult("E. Food & Beverage Vending"))

    line_5_total_tax_due: Decimal = Decimal("0.00")
    line_6_lawful_deductions: Decimal = Decimal("0.00")
    line_7_net_tax_due: Decimal = Decimal("0.00")
    line_8_est_tax_paid_credits: Decimal = Decimal("0.00")
    line_9_est_tax_due_current: Decimal = Decimal("0.00")
    line_10_amount_due: Decimal = Decimal("0.00")
    line_11_collection_allowance: Decimal = Decimal("0.00")
    line_12_penalty: Decimal = Decimal("0.00")
    line_13_interest: Decimal = Decimal("0.00")
    line_14_amount_due_with_return: Decimal = Decimal("0.00")

    # Back of form
    line_15a_exempt_over_5000: Decimal = Decimal("0.00")
    line_15b_other_not_subject_to_surtax: Decimal = Decimal("0.00")
    line_15c_different_surtax_rate_amount: Decimal = Decimal("0.00")
    line_15d_total_surtax_due: Decimal = Decimal("0.00")

    line_16_scholarship_credits: Decimal = Decimal("0.00")
    line_17_electricity_taxable: Decimal = Decimal("0.00")
    line_18_dyed_diesel_taxable: Decimal = Decimal("0.00")
    line_19_amusement_taxable: Decimal = Decimal("0.00")
    line_20_high_crime_credits: Decimal = Decimal("0.00")
    line_21_other_credits: Decimal = Decimal("0.00")

    # Meta
    late_filing: bool = False
    warnings: list = field(default_factory=list)

    # Shopify-vs-calculated comparison (populated when Shopify data is used)
    shopify_tax_collected: Decimal = Decimal("0.00")
    tax_gap: Decimal = Decimal("0.00")


# ---------- engine ----------
class FloridaSalesTaxEngine:
    def __init__(self, rules_path: str, business_config_path: str):
        # Load base rules, then merge any update files from rules/updates/
        self.rules = json.loads(Path(rules_path).read_text())
        self.applied_updates = self._apply_rule_updates(rules_path)
        self.business = json.loads(Path(business_config_path).read_text())
        self.state_rate = Decimal(str(self.rules["state_rates"]["general_sales_tax"]))
        self.electricity_rate = Decimal(str(self.rules["state_rates"]["electricity"]))

    def _apply_rule_updates(self, rules_path: str) -> list:
        """Merge any .json files from rules/updates/ on top of base rules.
        Files are applied in alphabetical order so newer dates override older.
        Files starting with '_' are treated as examples/disabled and skipped."""
        updates_dir = Path(rules_path).parent / "updates"
        applied = []
        if not updates_dir.exists():
            return applied
        for update_file in sorted(updates_dir.glob("*.json")):
            if update_file.name.startswith("_"):
                continue  # example/disabled files are skipped
            try:
                patch = json.loads(update_file.read_text())
                self._deep_merge(self.rules, patch)
                applied.append(update_file.name)
            except Exception as e:
                print(f"  ! Skipped rules update {update_file.name}: {e}")
        return applied

    @staticmethod
    def _deep_merge(base: dict, patch: dict) -> None:
        """Recursively merge patch into base (in place). Skips '_' meta keys
        at the top level of patch (they're just notes for the human)."""
        for k, v in patch.items():
            if k.startswith("_"):
                continue
            if isinstance(v, dict) and isinstance(base.get(k), dict):
                FloridaSalesTaxEngine._deep_merge(base[k], v)
            else:
                base[k] = v

    # ---- rates ----
    def county_surtax_rate(self, county: Optional[str] = None) -> Decimal:
        county = (county or self.business["business_info"]["county"]).replace(" ", "_")
        rate = self.rules["county_surtax_rates_2026"].get(county)
        if rate is None:
            return Decimal("0")
        return Decimal(str(rate))

    def total_rate(self, county: Optional[str] = None) -> Decimal:
        return self.state_rate + self.county_surtax_rate(county)

    # ---- Line A: regular sales ----
    def compute_line_a(
        self,
        transactions: list,
        county: Optional[str] = None,
    ) -> LineResult:
        """
        Works with BOTH the simple format and the rich Shopify format.

        Simple row keys (legacy):
          amount, exempt (bool), delivery_county, is_single_tpp_item, category

        Shopify row keys (new):
          gross_sales, discounts, returns, net_sales, exempt_amount,
          taxable_amount, shopify_tax_amount, delivery_county,
          destination_state, is_single_tpp_item, is_florida_sale
        """
        line = LineResult("A. Sales/Services/Electricity")
        gross = Decimal("0")
        exempt = Decimal("0")
        total_tax = Decimal("0")              # what the app CALCULATES is owed
        total_state_tax = Decimal("0")
        total_surtax = Decimal("0")
        shopify_tax_collected = Decimal("0")  # what was actually collected
        surtax_cap = Decimal(str(self.rules["surtax_rules"]["single_item_cap_tpp"]))

        amt_over_5000_exempt = Decimal("0")
        amt_different_surtax_rate = Decimal("0")
        amt_not_subject_to_surtax = Decimal("0")

        biz_surtax_rate = self.county_surtax_rate(county)

        for t in transactions:
            # -------- detect Shopify vs simple format --------
            is_shopify = "taxable_amount" in t and "gross_sales" in t

            if is_shopify:
                # Skip non-Florida sales (out-of-state buyers w/o nexus there)
                if not t.get("is_florida_sale", True):
                    continue
                g = Decimal(str(t.get("gross_sales", 0)))
                taxable = Decimal(str(t.get("taxable_amount", 0)))
                # DR-15 columns:
                #   Col 1 (gross)   = Shopify 'Gross sales on line items'
                #   Col 3 (taxable) = Shopify 'Taxable amount' — authoritative
                #   Col 2 (exempt)  = Col 1 − Col 3 (back-calculated)
                # This correctly absorbs discounts, exempt items, non-taxable
                # items, non-taxed items, AND return credits in one number.
                gross += g
                exempt += (g - taxable)
                shopify_tax_collected += Decimal(str(t.get("shopify_tax_amount", 0)))
                amount = taxable
            else:
                # Legacy path
                amount = Decimal(str(t["amount"]))
                gross += amount
                if t.get("exempt", False):
                    exempt += amount
                    continue

            category = t.get("category", "general")
            if category == "electricity":
                rate = self.electricity_rate
                state_tax = amount * rate
                surtax = amount * self.county_surtax_rate(t.get("delivery_county"))
                tx_tax = money(state_tax + surtax)
                total_tax += tx_tax
                total_state_tax += money(state_tax)
                total_surtax += money(surtax)
                continue

            # General TPP / services
            delivery_surtax = self.county_surtax_rate(t.get("delivery_county"))
            is_tpp = t.get("is_single_tpp_item", True)

            state_tax = amount * self.state_rate

            if is_tpp and amount > surtax_cap:
                surtax_base = surtax_cap
                amt_over_5000_exempt += amount - surtax_cap
            else:
                surtax_base = amount

            surtax = surtax_base * delivery_surtax

            if delivery_surtax != biz_surtax_rate:
                amt_different_surtax_rate += amount

            if delivery_surtax == 0:
                amt_not_subject_to_surtax += amount

            tx_tax = money(state_tax + surtax)
            total_tax += tx_tax
            total_state_tax += money(state_tax)
            total_surtax += money(surtax)

            # Stash computed values on the transaction for audit export
            if is_shopify:
                t["_calc_state_rate"] = float(self.state_rate)
                t["_calc_state_tax"] = float(money(state_tax))
                t["_calc_surtax_rate"] = float(delivery_surtax)
                t["_calc_surtax_base"] = float(surtax_base)
                t["_calc_surtax"] = float(money(surtax))
                t["_calc_total_tax"] = float(tx_tax)
                t["_calc_gap"] = float(
                    tx_tax - Decimal(str(t.get("shopify_tax_amount", 0)))
                )
                t["_calc_cap_applied"] = bool(is_tpp and amount > surtax_cap)
                t["_calc_exempt_portion"] = float(
                    Decimal(str(t.get("gross_sales", 0))) - amount
                )

        line.gross_sales = simple_round(gross)
        line.exempt_sales = simple_round(exempt)
        line.taxable_amount = simple_round(gross - exempt)
        line.tax_due = simple_round(total_tax)
        line.state_tax = simple_round(total_state_tax)
        line.surtax = simple_round(total_surtax)

        line._over_5000 = simple_round(amt_over_5000_exempt)
        line._different_surtax = simple_round(amt_different_surtax_rate)
        line._not_subject_to_surtax = simple_round(amt_not_subject_to_surtax)
        line._shopify_tax_collected = simple_round(shopify_tax_collected)
        line._tax_gap = simple_round(total_tax - shopify_tax_collected)

        return line

    # ---- Line B: use tax ----
    def compute_line_b(self, purchases: list, county: Optional[str] = None) -> LineResult:
        """
        purchases: list of dicts {amount, tax_paid_other_state (optional, decimal 0.xx rate or 0)}
        """
        line = LineResult("B. Taxable Purchases (Use Tax)")
        taxable = Decimal("0")
        total_tax = Decimal("0")
        total_state = Decimal("0")
        total_surtax = Decimal("0")
        surtax_rate = self.county_surtax_rate(county)

        for p in purchases:
            amount = Decimal(str(p["amount"]))
            taxable += amount
            tax_paid_other = Decimal(str(p.get("tax_paid_other_state", 0)))  # a rate like 0.04
            effective_state_rate = max(self.state_rate - tax_paid_other, Decimal("0"))
            state_tax = amount * effective_state_rate
            surtax = amount * surtax_rate  # simplified: no $5k cap tracking here
            total_state += state_tax
            total_surtax += surtax
            total_tax += state_tax + surtax

        line.gross_sales = Decimal("0.00")  # N/A for Line B
        line.exempt_sales = Decimal("0.00")
        line.taxable_amount = simple_round(taxable)
        line.state_tax = simple_round(total_state)
        line.surtax = simple_round(total_surtax)
        line.tax_due = simple_round(total_tax)
        return line

    # ---- Line D: transient rentals ----
    def compute_line_d(self, rentals: list, county: Optional[str] = None) -> LineResult:
        line = LineResult("D. Transient Rentals")
        gross = Decimal("0")
        exempt = Decimal("0")
        total_tax = Decimal("0")
        total_state = Decimal("0")
        total_surtax = Decimal("0")
        surtax_rate = self.county_surtax_rate(county)

        for r in rentals:
            amount = Decimal(str(r["amount"]))
            gross += amount
            if r.get("exempt", False):
                exempt += amount
                continue
            # No $5k cap for transient rentals
            state_tax = amount * self.state_rate
            surtax = amount * surtax_rate
            total_state += state_tax
            total_surtax += surtax
            total_tax += state_tax + surtax

        line.gross_sales = simple_round(gross)
        line.exempt_sales = simple_round(exempt)
        line.taxable_amount = simple_round(gross - exempt)
        line.state_tax = simple_round(total_state)
        line.surtax = simple_round(total_surtax)
        line.tax_due = simple_round(total_tax)
        return line

    # ---- Line E: food & beverage vending (divisor method) ----
    def compute_line_e(self, vending_receipts: list, county: Optional[str] = None) -> LineResult:
        """
        vending_receipts: list of {total_receipts, county (optional)}
        Uses food_and_beverage divisor table.
        """
        line = LineResult("E. Food & Beverage Vending")
        gross = Decimal("0")
        tax = Decimal("0")
        surtax_total = Decimal("0")

        for v in vending_receipts:
            receipts = Decimal(str(v["total_receipts"]))
            rate = self.total_rate(v.get("county", county))
            rate_key = f"{rate.quantize(Decimal('0.001'))}"  # '0.065'
            # match rates that are in the divisor table
            divisor_map = self.rules["vending_divisors"]["food_and_beverage"]
            divisor = None
            for k, v_ in divisor_map.items():
                if Decimal(k) == rate:
                    divisor = Decimal(str(v_))
                    break
            if divisor is None:
                # fallback to 1.0686 for 6.5%
                divisor = Decimal("1.0686")
            g = simple_round(receipts / divisor)
            t = simple_round(receipts - g)
            surtax_part = simple_round(g * self.county_surtax_rate(v.get("county", county)))
            gross += g
            tax += t
            surtax_total += surtax_part

        line.gross_sales = simple_round(gross)
        line.taxable_amount = simple_round(gross)
        line.tax_due = simple_round(tax)
        line.surtax = simple_round(surtax_total)
        line.state_tax = simple_round(tax - surtax_total)
        return line

    # ---- assemble full return ----
    def build_return(
        self,
        reporting_period: str,
        line_a_txns: Optional[list] = None,
        line_b_txns: Optional[list] = None,
        line_d_txns: Optional[list] = None,
        line_e_txns: Optional[list] = None,
        lawful_deductions: float = 0,
        est_tax_paid_last_month: float = 0,
        est_tax_due_current_month: float = 0,
        is_late: bool = False,
        scholarship_credits: float = 0,
        high_crime_credits: float = 0,
        other_credits: float = 0,
    ) -> DR15Return:
        ret = DR15Return(
            reporting_period=reporting_period,
            county=self.business["business_info"]["county"],
            surtax_rate=self.county_surtax_rate(),
            late_filing=is_late,
        )

        # build each active line
        active = self.business["active_lines"]
        if active.get("line_a_sales_services") and line_a_txns:
            ret.line_a = self.compute_line_a(line_a_txns)
        if active.get("line_b_use_tax_purchases") and line_b_txns:
            ret.line_b = self.compute_line_b(line_b_txns)
        if active.get("line_d_transient_rentals") and line_d_txns:
            ret.line_d = self.compute_line_d(line_d_txns)
        if active.get("line_e_food_beverage_vending") and line_e_txns:
            ret.line_e = self.compute_line_e(line_e_txns)

        # Line 5: total tax due (all column 4 totals)
        ret.line_5_total_tax_due = simple_round(
            ret.line_a.tax_due + ret.line_b.tax_due + ret.line_c.tax_due
            + ret.line_d.tax_due + ret.line_e.tax_due
        )

        # Line 6: lawful deductions capped at Line 5
        ld = Decimal(str(lawful_deductions)) + Decimal(str(scholarship_credits))
        if ld > ret.line_5_total_tax_due:
            ret.warnings.append(
                f"Lawful deductions (${ld}) capped at Line 5 tax due (${ret.line_5_total_tax_due}). "
                f"Remainder ${ld - ret.line_5_total_tax_due} may be carried to next return."
            )
            ld = ret.line_5_total_tax_due
        ret.line_6_lawful_deductions = simple_round(ld)
        ret.line_16_scholarship_credits = simple_round(scholarship_credits)

        # Line 7: net tax due
        ret.line_7_net_tax_due = simple_round(ret.line_5_total_tax_due - ret.line_6_lawful_deductions)

        # Line 8: est tax paid + credit memos, capped at Line 7
        l8 = Decimal(str(est_tax_paid_last_month)) + Decimal(str(high_crime_credits)) + Decimal(str(other_credits))
        if l8 > ret.line_7_net_tax_due:
            ret.warnings.append(
                f"Line 8 credits (${l8}) capped at Line 7 net tax due (${ret.line_7_net_tax_due})."
            )
            l8 = ret.line_7_net_tax_due
        ret.line_8_est_tax_paid_credits = simple_round(l8)
        ret.line_20_high_crime_credits = simple_round(high_crime_credits)
        ret.line_21_other_credits = simple_round(other_credits)

        # Line 9: estimated tax due
        ret.line_9_est_tax_due_current = simple_round(est_tax_due_current_month)

        # Line 10: amount due (cannot be negative)
        l10 = ret.line_7_net_tax_due - ret.line_8_est_tax_paid_credits + ret.line_9_est_tax_due_current
        ret.line_10_amount_due = simple_round(max(l10, Decimal("0")))

        # Line 11: collection allowance (only if e-file + e-pay + timely)
        prefs = self.business["filing_preferences"]
        if prefs["files_electronically"] and prefs["pays_electronically"] and not is_late:
            ca_cfg = self.rules["collection_allowance"]
            rate = Decimal(str(ca_cfg["rate"]))
            cap_amt = Decimal(str(ca_cfg["applies_to_first"]))
            max_ca = Decimal(str(ca_cfg["max_allowance"]))
            base = min(ret.line_10_amount_due, cap_amt)
            allowance = min(base * rate, max_ca)
            if prefs.get("donate_collection_allowance_to_education"):
                ret.line_11_collection_allowance = Decimal("0.00")
                ret.warnings.append("Collection allowance donated to Education Trust Fund (Line 11 blank).")
            else:
                ret.line_11_collection_allowance = simple_round(allowance)

        # Lines 12-13: penalty & interest if late
        if is_late:
            pen_pct = Decimal(str(self.rules["penalty_and_interest"]["late_penalty_percent"]))
            min_pen = Decimal(str(self.rules["penalty_and_interest"]["minimum_late_penalty"]))
            pen = max(ret.line_10_amount_due * pen_pct, min_pen)
            ret.line_12_penalty = simple_round(pen)
            # Interest: placeholder — needs days late to compute exactly
            ret.warnings.append(
                "Return marked late: Line 13 interest requires days-late calculation. "
                "Use floating daily rate from floridarevenue.com/taxes/rates."
            )

        # Line 14: amount due with return
        if is_late:
            ret.line_14_amount_due_with_return = simple_round(
                ret.line_10_amount_due + ret.line_12_penalty + ret.line_13_interest
            )
        else:
            ret.line_14_amount_due_with_return = simple_round(
                ret.line_10_amount_due - ret.line_11_collection_allowance
            )

        # Back of form: 15(a)-(d)
        ret.line_15a_exempt_over_5000 = getattr(ret.line_a, "_over_5000", Decimal("0.00"))
        ret.line_15b_other_not_subject_to_surtax = getattr(ret.line_a, "_not_subject_to_surtax", Decimal("0.00"))
        ret.line_15c_different_surtax_rate_amount = getattr(ret.line_a, "_different_surtax", Decimal("0.00"))
        ret.line_15d_total_surtax_due = simple_round(
            ret.line_a.surtax + ret.line_b.surtax + ret.line_c.surtax
            + ret.line_d.surtax + ret.line_e.surtax
        )

        # Shopify collection-gap analysis (only populated for Shopify inputs)
        shopify_collected = getattr(ret.line_a, "_shopify_tax_collected", Decimal("0.00"))
        tax_gap = getattr(ret.line_a, "_tax_gap", Decimal("0.00"))
        if shopify_collected > 0:
            ret.shopify_tax_collected = shopify_collected
            ret.tax_gap = tax_gap
            if abs(tax_gap) >= Decimal("1.00"):
                if tax_gap > 0:
                    ret.warnings.append(
                        f"Under-collection detected: Shopify collected ${shopify_collected} "
                        f"but ${ret.line_a.tax_due} is owed (${tax_gap} short). "
                        f"You must remit the full calculated amount and absorb the shortfall. "
                        f"Check that Shopify is configured to charge the county surtax."
                    )
                else:
                    ret.warnings.append(
                        f"Over-collection detected: Shopify collected ${shopify_collected} "
                        f"but only ${ret.line_a.tax_due} is owed (${-tax_gap} overage). "
                        f"Remit the full amount collected (per Florida law, you can't keep excess)."
                    )

        return ret


# ---------- helper for pretty dict dump ----------
def return_to_dict(r: DR15Return) -> dict:
    def conv(v):
        if isinstance(v, Decimal):
            return float(v)
        if isinstance(v, LineResult):
            return {k: float(val) if isinstance(val, Decimal) else val
                    for k, val in asdict(v).items() if not k.startswith("_")}
        return v
    d = {}
    for k, v in asdict(r).items():
        d[k] = conv(v) if not isinstance(v, dict) else v
    return d
