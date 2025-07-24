"""
Reusable Excel worksheet tools for LangChain and app integration.
"""
import pandas as pd
from typing import List, Optional, Dict, Any

def read_worksheet(excel_file: str, sheet_name: str) -> pd.DataFrame:
    """Read a worksheet from an Excel file."""
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
        return df
    except Exception as e:
        raise RuntimeError(f"Failed to read worksheet '{sheet_name}': {e}")

def filter_data(df: pd.DataFrame, conditions: str) -> pd.DataFrame:
    """Filter DataFrame using a pandas query string."""
    try:
        filtered = df.query(conditions)
        return filtered
    except Exception as e:
        raise RuntimeError(f"Failed to filter data: {e}")

def aggregate_data(df: pd.DataFrame, group_by: List[str], metrics: Dict[str, str]) -> pd.DataFrame:
    """Aggregate data using groupby and specified metrics.
    metrics: e.g. { 'Sales': 'sum', 'Quantity': 'mean' }
    """
    try:
        agg_df = df.groupby(group_by).agg(metrics).reset_index()
        return agg_df
    except Exception as e:
        raise RuntimeError(f"Failed to aggregate data: {e}")

def sort_data(df: pd.DataFrame, by: List[str], ascending: Optional[List[bool]] = None) -> pd.DataFrame:
    """Sort DataFrame by columns."""
    try:
        if ascending is None:
            ascending = [True] * len(by)
        sorted_df = df.sort_values(by=by, ascending=ascending)
        return sorted_df
    except Exception as e:
        raise RuntimeError(f"Failed to sort data: {e}")

def pivot_table(df: pd.DataFrame, index: List[str], columns: List[str], values: List[str], aggfunc: str = 'sum') -> pd.DataFrame:
    """Create a pivot table from DataFrame."""
    try:
        pivot = pd.pivot_table(df, index=index, columns=columns, values=values, aggfunc=aggfunc, fill_value=0)
        return pivot.reset_index()
    except Exception as e:
        raise RuntimeError(f"Failed to create pivot table: {e}")

# (Optional) Advanced tool stubs for next phases
def merge_worksheets(*args, **kwargs):
    raise NotImplementedError("merge_worksheets is not implemented yet.")
def data_validation(*args, **kwargs):
    raise NotImplementedError("data_validation is not implemented yet.")
def formula_evaluation(*args, **kwargs):
    raise NotImplementedError("formula_evaluation is not implemented yet.")
def chart_generation(*args, **kwargs):
    raise NotImplementedError("chart_generation is not implemented yet.")
