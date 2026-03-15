import warnings
from typing import Optional
import pandas as pd

# Claude AI assistant helped me with this function

# Lookup table: maps a frequency name to its pandas alias and a rank.
# Higher rank = broader time period (monthly=1, quarterly=2, yearly=3).
_FREQ_MAP = {
    "monthly":   ("M",  1),
    "quarterly": ("Q",  2),
    "yearly":    ("A",  3),
}

def resample_series(df: pd.DataFrame, freq: str, agg: str = "mean") -> pd.DataFrame:
    """
    Convert a FRED time series to a broader time frequency.

    For example, turn monthly data into quarterly or yearly data.
    If you ask for a finer frequency than the data already has
    (e.g. monthly from a yearly dataset), the data is returned
    as-is with a warning — we never make up data points.
    """
    freq = freq.lower()

    # Make sure the user picked a valid frequency
    if freq not in _FREQ_MAP:
        raise ValueError(
            f"Unsupported frequency '{freq}'. "
            f"Choose from: {list(_FREQ_MAP.keys())}"
        )

    # Make sure the user picked a valid aggregation method
    valid_aggs = {"mean", "sum", "last", "first"}
    if agg not in valid_aggs:
        raise ValueError(
            f"Unsupported aggregation '{agg}'. "
            f"Choose from: {sorted(valid_aggs)}"
        )

    target_alias, target_rank = _FREQ_MAP[freq]

    # Work on a copy so the original DataFrame is never modified
    result = df.copy()
    result = result.set_index(pd.DatetimeIndex(result["date"]))
    result = result[["value"]]

    # Figure out how frequently the incoming data is recorded
    native_alias = pd.infer_freq(result.index)
    native_rank = _infer_rank(native_alias)

    # return a warning if the requested frequency is finer than the data we have.
    # For example, we can't turn yearly data into monthly data
    if native_rank is not None and target_rank < native_rank:
        warnings.warn(
            f"The dataset's native frequency ({native_alias}) is coarser than "
            f"the requested frequency ('{freq}'). Returning data at its native "
            f"frequency to avoid fabricating values not present in FRED.",
            UserWarning,
            stacklevel=2,
        )
        result = result.reset_index().rename(columns={"index": "date"})
        result["date"] = result["date"].dt.normalize()
        return result[["date", "value"]].reset_index(drop=True)

    # Combine rows into the target time period using the chosen aggregation
    agg_funcs = {
        "mean":  result["value"].resample(target_alias).mean,
        "sum":   result["value"].resample(target_alias).sum,
        "last":  result["value"].resample(target_alias).last,
        "first": result["value"].resample(target_alias).first,
    }
    resampled = agg_funcs[agg]()

    # Remove any empty time periods
    resampled = resampled.dropna()

    #reset the index and normalize timestamps to midnight
    result = resampled.reset_index()
    result.columns = ["date", "value"]
    result["date"] = result["date"].dt.normalize()

    return result


# helpers
def _infer_rank(alias: Optional[str]) -> Optional[int]:
    """
    Turn a pandas frequency alias (like "QS-OCT" or "MS") into a rank number
    so we can compare how broad two frequencies are.

    Rank scale:
      0 = very fine  (daily, weekly, hourly, etc.)
      1 = monthly
      2 = quarterly
      3 = yearly

    Returns None if the alias is not recognized.
    """
    if alias is None:
        return None

    alias_upper = alias.upper()

    # Check prefix since pandas generates many variants of the same frequency
    if any(alias_upper.startswith(p) for p in ("Y", "A")):
        return 3  # yearly
    if alias_upper.startswith("Q"):
        return 2  # quarterly
    if alias_upper.startswith("M"):
        return 1  # monthly
    if any(alias_upper.startswith(p) for p in ("W", "D", "H", "T", "S", "B")):
        return 0  # weekly, daily, or finer

    return None  # unrecognized alias