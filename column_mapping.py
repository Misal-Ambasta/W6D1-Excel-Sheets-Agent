"""
Column Name Intelligence: Fuzzy matching, synonym mapping, normalization, and LLM-assisted suggestions.
"""
import re
from typing import List, Dict, Tuple, Optional
from rapidfuzz import fuzz, process
from llm_utils import run_gemini_query

# Business synonym dictionary
SYNONYM_DICT = {
    'qty': 'quantity',
    'quantity': 'qty',
    'amt': 'amount',
    'amount': 'amt',
    'cust': 'customer',
    'customer': 'cust',
    'prod': 'product',
    'product': 'prod',
    'desc': 'description',
    'description': 'desc',
    # Add more as needed
}

def normalize_header(header: str) -> str:
    """Normalize a header: lowercase, remove special chars, replace spaces with _"""
    header = header.lower()
    header = re.sub(r'[^a-z0-9_ ]', '', header)
    header = header.replace(' ', '_')
    return header

def normalize_headers(headers: List[str]) -> List[str]:
    return [normalize_header(h) for h in headers]

def fuzzy_match_column(query: str, candidates: List[str], threshold: int = 80) -> Optional[Tuple[str, int]]:
    """Return best fuzzy match for query in candidates above threshold, else None."""
    result = process.extractOne(query, candidates, scorer=fuzz.ratio)
    if result and result[1] >= threshold:
        return result
    return None

def map_with_synonyms(header: str, candidates: List[str]) -> Optional[str]:
    """Try to map header using synonyms."""
    norm_header = normalize_header(header)
    # Direct synonym
    if norm_header in SYNONYM_DICT:
        synonym = SYNONYM_DICT[norm_header]
        match = fuzzy_match_column(synonym, candidates, threshold=70)
        if match:
            return match[0]
    return None

def suggest_column_mapping_with_llm(user_header: str, candidates: List[str]) -> Optional[str]:
    """Use LLM to suggest a mapping for an ambiguous column name."""
    prompt = (
        f"Given the ambiguous column name '{user_header}' and the available columns: {candidates}, "
        "suggest the most likely correct column name. Return only the best match or None."
    )
    suggestion = run_gemini_query('filter', prompt, candidates)
    if suggestion:
        # Clean up LLM response
        suggestion = suggestion.strip().strip('"\'')
        if suggestion in candidates:
            return suggestion
    return None

def map_column(user_header: str, candidates: List[str]) -> Tuple[Optional[str], str]:
    """
    Map a user-supplied column name to the best candidate using normalization, synonyms, fuzzy, and LLM.
    Returns (best_match, method_used)
    """
    # Normalize all
    norm_candidates = normalize_headers(candidates)
    norm_user_header = normalize_header(user_header)
    # 1. Exact match
    if norm_user_header in norm_candidates:
        idx = norm_candidates.index(norm_user_header)
        return candidates[idx], 'exact'
    # 2. Synonym match
    syn_match = map_with_synonyms(user_header, candidates)
    if syn_match:
        return syn_match, 'synonym'
    # 3. Fuzzy match
    fuzzy = fuzzy_match_column(user_header, candidates)
    if fuzzy:
        return fuzzy[0], 'fuzzy'
    # 4. LLM suggestion
    llm = suggest_column_mapping_with_llm(user_header, candidates)
    if llm:
        return llm, 'llm'
    return None, 'none'
