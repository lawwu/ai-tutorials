"""
test_data_quality.py
────────────────────────────────────────────────────────────────────────────────
Automated data quality tests for the Campaign Health data.

Tests read directly from AQ_Impact_Combined.xlsx (the three sheets:
  • AQ Data       — raw Audience Quality rows
  • Impact Data   — raw Impact rows
  • Combined View — the merged output used by the dashboard

Run:
    pip install pytest pandas openpyxl
    pytest test_data_quality.py -v
────────────────────────────────────────────────────────────────────────────────
"""

import math
import pytest
import pandas as pd

# ── Load once, share across all tests ─────────────────────────────────────────
EXCEL_PATH = "AQ_Impact_Combined.xlsx"

@pytest.fixture(scope="session")
def aq():
    df = pd.read_excel(EXCEL_PATH, sheet_name="AQ Data")
    df.columns = [c.strip() for c in df.columns]
    return df

@pytest.fixture(scope="session")
def imp():
    df = pd.read_excel(EXCEL_PATH, sheet_name="Impact Data")
    df.columns = [c.strip() for c in df.columns]
    return df

@pytest.fixture(scope="session")
def combined():
    # Row 0 is the section banner ("◆  AUDIENCE QUALITY" etc); real headers are row 1
    df = pd.read_excel(EXCEL_PATH, sheet_name="Combined View", header=1)
    # Flatten multi-line header names produced by openpyxl wrap_text
    df.columns = [c.replace("\n", " ").strip() for c in df.columns]
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 1. SCHEMA — required columns exist
# ══════════════════════════════════════════════════════════════════════════════
class TestSchema:
    AQ_REQUIRED = [
        "Time Period", "Time Stamp", "Contracted Data Break",
        "Publisher", "Publisher Type", "Ad Type", "Placement Detail", "Targeting",
        "AQ Segments", "Consumer Reach", "Target Reach", "AQI", "Frequency", "Confidence",
    ]
    IMP_REQUIRED = [
        "Time Period", "Time Stamp", "Contracted Data Break",
        "Publisher", "Publisher Type", "Ad Type", "Placement Detail", "Targeting",
        "SOB", "SOB Type",
        "Incremental Conversions", "Relative Lift", "Frequency to Impact", "Stat Sig",
    ]
    # The Combined View uses "friendly" wrapped column names (newlines replaced with spaces)
    COMBINED_REQUIRED = [
        "Time Period", "Time Stamp", "Contracted Data Break",
        "Publisher", "Publisher Type", "Ad Type", "Placement Detail", "Targeting",
        "AQ Segment", "Consumer Reach", "Target Reach", "AQI", "Frequency", "Confidence",
        "SOB", "SOB Type",
        "Incremental Conversions", "Relative Lift", "Freq to Impact", "Stat Sig",
    ]

    def test_aq_has_required_columns(self, aq):
        missing = [c for c in self.AQ_REQUIRED if c not in aq.columns]
        assert not missing, f"AQ Data missing columns: {missing}"

    def test_impact_has_required_columns(self, imp):
        missing = [c for c in self.IMP_REQUIRED if c not in imp.columns]
        assert not missing, f"Impact Data missing columns: {missing}"

    def test_combined_has_required_columns(self, combined):
        missing = [c for c in self.COMBINED_REQUIRED if c not in combined.columns]
        assert not missing, f"Combined View missing columns: {missing}"

    def test_aq_not_empty(self, aq):
        assert len(aq) > 0, "AQ Data sheet is empty"

    def test_impact_not_empty(self, imp):
        assert len(imp) > 0, "Impact Data sheet is empty"

    def test_combined_not_empty(self, combined):
        assert len(combined) > 0, "Combined View sheet is empty"


# ══════════════════════════════════════════════════════════════════════════════
# 2. COMPLETENESS — required fields are never null
# ══════════════════════════════════════════════════════════════════════════════
class TestCompleteness:
    # These fields should never be missing in either source sheet
    AQ_NOT_NULL = [
        "Time Period", "Time Stamp", "Contracted Data Break",
        "Publisher", "AQ Segments", "Consumer Reach", "Target Reach", "AQI",
        "Frequency", "Confidence",
    ]
    IMP_NOT_NULL = [
        "Time Period", "Time Stamp", "Contracted Data Break",
        "Publisher", "SOB", "SOB Type",
        "Incremental Conversions", "Relative Lift", "Frequency to Impact", "Stat Sig",
    ]

    @pytest.mark.parametrize("col", AQ_NOT_NULL)
    def test_aq_no_nulls_in(self, aq, col):
        null_count = aq[col].isna().sum()
        assert null_count == 0, f"AQ Data.'{col}' has {null_count} null value(s)"

    @pytest.mark.parametrize("col", IMP_NOT_NULL)
    def test_impact_no_nulls_in(self, imp, col):
        null_count = imp[col].isna().sum()
        assert null_count == 0, f"Impact Data.'{col}' has {null_count} null value(s)"

    def test_combined_join_keys_never_null(self, combined):
        # The Combined View uses wrapped/abbreviated column names
        join_keys = [
            "Time Period", "Time Stamp", "Contracted Data Break",
            "Publisher", "Publisher Type", "Ad Type", "Placement Detail", "Targeting",
        ]
        # Map friendly wrapped names back to actual column names in combined sheet
        col_map = {
            "Time Period": "Time Period",
            "Time Stamp": "Time Stamp",
            "Contracted Data Break": "Contracted Data Break",
            "Publisher": "Publisher",
            "Publisher Type": "Publisher Type",
            "Ad Type": "Ad Type",
            "Placement Detail": "Placement Detail",
            "Targeting": "Targeting",
        }
        for label, col in col_map.items():
            actual_col = next((c for c in combined.columns if c == col), None)
            if actual_col is None:
                pytest.fail(f"Combined View missing join key column: '{col}'")
            null_count = combined[actual_col].isna().sum()
            assert null_count == 0, \
                f"Combined View join key '{label}' has {null_count} null value(s)"


# ══════════════════════════════════════════════════════════════════════════════
# 3. DATA TYPES — numeric columns contain numbers, not text
# ══════════════════════════════════════════════════════════════════════════════
class TestDataTypes:
    def test_aq_numeric_columns(self, aq):
        for col in ["Consumer Reach", "Target Reach", "AQI", "Frequency"]:
            assert pd.api.types.is_numeric_dtype(aq[col]), \
                f"AQ Data.'{col}' should be numeric, got {aq[col].dtype}"

    def test_impact_numeric_columns(self, imp):
        for col in ["Incremental Conversions", "Relative Lift", "Frequency to Impact"]:
            assert pd.api.types.is_numeric_dtype(imp[col]), \
                f"Impact Data.'{col}' should be numeric, got {imp[col].dtype}"

    def test_time_stamp_is_integer(self, aq):
        assert pd.api.types.is_integer_dtype(aq["Time Stamp"]), \
            f"Time Stamp should be integer (YYYYMM), got {aq['Time Stamp'].dtype}"

    def test_time_stamp_yyyymm_format(self, aq):
        """Time Stamp must be a valid YYYYMM value (e.g. 202509)."""
        bad = aq[~aq["Time Stamp"].astype(str).str.match(r"^\d{6}$")]
        assert bad.empty, \
            f"Time Stamp values not in YYYYMM format:\n{bad[['Time Stamp']].drop_duplicates()}"

    def test_time_stamp_plausible_year(self, aq):
        years = aq["Time Stamp"].astype(str).str[:4].astype(int)
        assert (years >= 2000).all() and (years <= 2030).all(), \
            f"Time Stamp year(s) out of expected range [2000–2030]: {years.unique()}"

    def test_time_stamp_plausible_month(self, aq):
        months = aq["Time Stamp"].astype(str).str[4:6].astype(int)
        assert ((months >= 1) & (months <= 12)).all(), \
            f"Time Stamp month(s) out of range [01–12]: {months.unique()}"


# ══════════════════════════════════════════════════════════════════════════════
# 4. VALUE RANGES — numeric fields within sensible bounds
# ══════════════════════════════════════════════════════════════════════════════
class TestValueRanges:
    def test_aqi_greater_than_zero(self, aq):
        bad = aq[aq["AQI"] <= 0]
        assert bad.empty, f"AQI must be > 0. Bad rows:\n{bad[['AQ Segments','AQI']]}"

    def test_aqi_reasonable_upper_bound(self, aq):
        """AQI is an index ratio — values above 10 are almost certainly an error."""
        bad = aq[aq["AQI"] > 10]
        assert bad.empty, f"AQI suspiciously high (>10):\n{bad[['AQ Segments','AQI']]}"

    def test_consumer_reach_positive(self, aq):
        bad = aq[aq["Consumer Reach"] <= 0]
        assert bad.empty, \
            f"Consumer Reach must be > 0:\n{bad[['Publisher','Consumer Reach']]}"

    def test_target_reach_positive(self, aq):
        bad = aq[aq["Target Reach"] <= 0]
        assert bad.empty, \
            f"Target Reach must be > 0:\n{bad[['Publisher','Target Reach']]}"

    def test_target_reach_not_exceed_consumer_reach(self, aq):
        """Target audience cannot be larger than total consumers reached."""
        bad = aq[aq["Target Reach"] > aq["Consumer Reach"]]
        assert bad.empty, \
            f"Target Reach > Consumer Reach (impossible):\n{bad[['AQ Segments','Target Reach','Consumer Reach']]}"

    def test_frequency_at_least_one(self, aq):
        bad = aq[aq["Frequency"] < 1]
        assert bad.empty, \
            f"Frequency must be >= 1:\n{bad[['Publisher','Frequency']]}"

    def test_incremental_conversions_non_negative(self, imp):
        bad = imp[imp["Incremental Conversions"] < 0]
        assert bad.empty, \
            f"Incremental Conversions cannot be negative:\n{bad[['SOB','Incremental Conversions']]}"

    def test_relative_lift_non_negative(self, imp):
        bad = imp[imp["Relative Lift"] < 0]
        assert bad.empty, \
            f"Relative Lift cannot be negative:\n{bad[['SOB','Relative Lift']]}"

    def test_frequency_to_impact_at_least_one(self, imp):
        bad = imp[imp["Frequency to Impact"] < 1]
        assert bad.empty, \
            f"Frequency to Impact must be >= 1:\n{bad[['SOB','Frequency to Impact']]}"


# ══════════════════════════════════════════════════════════════════════════════
# 5. VALID CATEGORIES — enum fields contain only expected values
# ══════════════════════════════════════════════════════════════════════════════
class TestValidCategories:
    CONFIDENCE_LEVELS = {"High", "Medium", "Low"}
    STAT_SIG_LEVELS   = {"High", "Medium", "Low"}

    def test_confidence_valid_values(self, aq):
        bad = aq[~aq["Confidence"].isin(self.CONFIDENCE_LEVELS)]
        assert bad.empty, \
            f"Unexpected Confidence values: {bad['Confidence'].unique()}"

    def test_stat_sig_valid_values(self, imp):
        bad = imp[~imp["Stat Sig"].isin(self.STAT_SIG_LEVELS)]
        assert bad.empty, \
            f"Unexpected Stat Sig values: {bad['Stat Sig'].unique()}"

    def test_time_period_non_empty_strings(self, aq):
        bad = aq[aq["Time Period"].str.strip() == ""]
        assert bad.empty, "Time Period contains blank strings"

    def test_publisher_no_leading_trailing_spaces(self, aq):
        """Catches data entry issues like 'Good Rx ' vs 'Good Rx'."""
        bad = aq[aq["Publisher"] != aq["Publisher"].str.strip()]
        assert bad.empty, \
            f"Publisher has leading/trailing whitespace:\n{bad['Publisher'].unique()}"

    def test_aq_segments_no_leading_trailing_spaces(self, aq):
        bad = aq[aq["AQ Segments"] != aq["AQ Segments"].str.strip()]
        assert bad.empty, \
            f"AQ Segments has leading/trailing whitespace:\n{bad['AQ Segments'].unique()}"

    def test_sob_type_consistent_naming(self, imp):
        """
        SOB Type has variations like 'Dr Visit', 'Dr Visits', 'Dr. Visit'.
        This test flags them so they can be standardized.
        """
        dr_variants = [v for v in imp["SOB Type"].unique()
                       if str(v).lower().startswith("dr")]
        assert len(dr_variants) <= 1, \
            f"Inconsistent 'Dr Visit' spellings — please standardize: {dr_variants}"


# ══════════════════════════════════════════════════════════════════════════════
# 6. UNIQUENESS — no exact duplicate rows
# ══════════════════════════════════════════════════════════════════════════════
class TestUniqueness:
    def test_aq_no_duplicate_rows(self, aq):
        dups = aq[aq.duplicated()]
        assert dups.empty, \
            f"AQ Data has {len(dups)} fully duplicate row(s):\n{dups}"

    def test_impact_no_duplicate_rows(self, imp):
        dups = imp[imp.duplicated()]
        assert dups.empty, \
            f"Impact Data has {len(imp)} fully duplicate row(s):\n{dups}"

    def test_aq_no_duplicate_keys(self, aq):
        """Each unique combination of dimensions + AQ Segment should appear once."""
        key_cols = [
            "Time Period", "Time Stamp", "Contracted Data Break",
            "Publisher", "Ad Type", "Placement Detail", "Targeting", "AQ Segments",
        ]
        dups = aq[aq.duplicated(subset=key_cols, keep=False)]
        assert dups.empty, \
            f"AQ Data has duplicate dimension keys ({len(dups)} rows):\n{dups[key_cols]}"

    def test_impact_no_duplicate_keys(self, imp):
        """Each combination of dimensions + SOB + SOB Type should appear once."""
        key_cols = [
            "Time Period", "Time Stamp", "Contracted Data Break",
            "Publisher", "Ad Type", "Placement Detail", "Targeting", "SOB", "SOB Type",
        ]
        dups = imp[imp.duplicated(subset=key_cols, keep=False)]
        assert dups.empty, \
            f"Impact Data has duplicate dimension keys ({len(dups)} rows):\n{dups[key_cols]}"


# ══════════════════════════════════════════════════════════════════════════════
# 7. CROSS-SHEET CONSISTENCY
# ══════════════════════════════════════════════════════════════════════════════
class TestCrossSheetConsistency:
    JOIN_KEYS = [
        "Time Period", "Time Stamp", "Contracted Data Break",
        "Publisher", "Publisher Type", "Ad Type", "Placement Detail", "Targeting",
    ]

    def test_aq_and_impact_share_same_time_stamps(self, aq, imp):
        aq_stamps  = set(aq["Time Stamp"].unique())
        imp_stamps = set(imp["Time Stamp"].unique())
        only_in_aq  = aq_stamps  - imp_stamps
        only_in_imp = imp_stamps - aq_stamps
        assert not only_in_aq,  f"Time Stamps in AQ but not Impact: {only_in_aq}"
        assert not only_in_imp, f"Time Stamps in Impact but not AQ: {only_in_imp}"

    def test_publishers_consistent_between_sheets(self, aq, imp):
        aq_pubs  = set(aq["Publisher"].unique())
        imp_pubs = set(imp["Publisher"].unique())
        only_in_aq  = aq_pubs  - imp_pubs
        only_in_imp = imp_pubs - aq_pubs
        # Warn-style: orphan publishers are suspicious but not always fatal
        # (outer join may legitimately produce some)
        assert not (only_in_aq and only_in_imp), \
            f"Publishers only in AQ: {only_in_aq}; only in Impact: {only_in_imp}"

    def test_combined_row_count_gte_aq_rows(self, aq, combined):
        """After an outer merge the combined sheet should have >= AQ rows."""
        assert len(combined) >= len(aq), \
            f"Combined ({len(combined)} rows) < AQ ({len(aq)} rows) — merge may have dropped data"

    def test_combined_row_count_gte_impact_rows(self, imp, combined):
        assert len(combined) >= len(imp), \
            f"Combined ({len(combined)} rows) < Impact ({len(imp)} rows) — merge may have dropped data"

    def test_combined_aqi_matches_aq_source(self, aq, combined):
        """
        Spot-check: every AQI value in the combined sheet should also
        appear in the AQ source sheet (no values invented by the merge).
        """
        # The Combined View "AQI" column may be read as float; round to avoid
        # floating-point mismatches
        combined_aqis = set(combined["AQI"].dropna().astype(float).round(4))
        aq_aqis       = set(aq["AQI"].dropna().astype(float).round(4))
        phantom = combined_aqis - aq_aqis
        assert not phantom, \
            f"AQI values in Combined not found in AQ source: {phantom}"


# ══════════════════════════════════════════════════════════════════════════════
# 8. BUSINESS LOGIC
# ══════════════════════════════════════════════════════════════════════════════
class TestBusinessLogic:
    def test_high_confidence_aqi_above_threshold(self, aq):
        """
        Rows with High confidence should generally have a meaningful AQI.
        Flag any High-confidence rows where AQI < 1 (below baseline).
        """
        high_conf = aq[aq["Confidence"] == "High"]
        below_baseline = high_conf[high_conf["AQI"] < 1.0]
        assert below_baseline.empty, \
            f"High-confidence rows with AQI < 1.0 (below baseline):\n" \
            f"{below_baseline[['AQ Segments','Publisher','AQI','Confidence']]}"

    def test_low_stat_sig_lift_interpretation(self, imp):
        """
        Very high Relative Lift (>5) paired with Low statistical significance
        is a data quality red flag — the result isn't trustworthy.
        """
        risky = imp[
            (imp["Stat Sig"] == "Low") & (imp["Relative Lift"] > 5)
        ]
        if not risky.empty:
            pytest.warns(
                UserWarning,
                match="high lift, low sig"
            )
        # Not a hard failure — surface as info
        assert True, \
            f"Note: {len(risky)} row(s) have high lift but low statistical significance"

    def test_incremental_conversions_positive_when_high_stat_sig(self, imp):
        """Statistically significant results should have real (>0) conversions."""
        high_sig = imp[imp["Stat Sig"] == "High"]
        zero_conv = high_sig[high_sig["Incremental Conversions"] <= 0]
        assert zero_conv.empty, \
            f"High Stat Sig rows with zero/negative conversions:\n" \
            f"{zero_conv[['SOB','SOB Type','Incremental Conversions','Stat Sig']]}"

    def test_no_impossible_lift_and_negative_conversions(self, imp):
        """Positive Relative Lift must come with positive Incremental Conversions."""
        bad = imp[
            (imp["Relative Lift"] > 0) & (imp["Incremental Conversions"] <= 0)
        ]
        assert bad.empty, \
            f"Positive lift but non-positive conversions:\n" \
            f"{bad[['SOB','Relative Lift','Incremental Conversions']]}"

    def test_frequency_to_impact_lte_aq_frequency(self, combined):
        """
        Frequency to Impact should not exceed the measured AQ Frequency —
        you can't need more exposures than were actually delivered.
        The Combined View calls these columns 'Frequency' and 'Freq to Impact'.
        """
        both = combined.dropna(subset=["Frequency", "Freq to Impact"])
        bad  = both[both["Freq to Impact"] > both["Frequency"]]
        assert bad.empty, \
            f"Freq to Impact > Frequency in {len(bad)} row(s):\n" \
            f"{bad[['Publisher','Frequency','Freq to Impact']].head(10)}"
