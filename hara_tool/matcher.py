"""
Fuzzy matching module for HARA Tool
"""

import pandas as pd
from fuzzywuzzy import fuzz
from typing import Tuple, Optional, Dict, Any


class FuzzyMatcher:
    """Handles fuzzy string matching for scenario and hazard matching"""
    
    def __init__(self, 
                 enabled: bool = True,
                 threshold: int = 80,
                 algorithm: str = 'Ratio (Default)',
                 case_sensitive: bool = False,
                 strip_whitespace: bool = True):
        """
        Initialize the fuzzy matcher with settings.
        
        Args:
            enabled: Whether fuzzy matching is enabled
            threshold: Minimum score to consider a match (0-100)
            algorithm: Matching algorithm to use
            case_sensitive: Whether to consider case in matching
            strip_whitespace: Whether to strip extra whitespace
        """
        self.enabled = enabled
        self.threshold = threshold
        self.algorithm = algorithm
        self.case_sensitive = case_sensitive
        self.strip_whitespace = strip_whitespace
        
    def prepare_string(self, text: Any) -> str:
        """
        Prepare string for comparison based on settings.
        
        Args:
            text: Input text (can be any type)
            
        Returns:
            Prepared string for comparison
        """
        if pd.isna(text) or text is None:
            return ""
        
        text = str(text)
        
        if self.strip_whitespace:
            text = ' '.join(text.split())  # Remove extra whitespace
            
        if not self.case_sensitive:
            text = text.lower()
            
        return text.strip()
    
    def calculate_score(self, str1: str, str2: str) -> int:
        """
        Calculate fuzzy match score between two strings.
        
        Args:
            str1: First string
            str2: Second string
            
        Returns:
            Match score (0-100)
        """
        str1 = self.prepare_string(str1)
        str2 = self.prepare_string(str2)
        
        if not str1 or not str2:
            return 0
        
        # Select algorithm based on setting
        if 'Partial' in self.algorithm:
            return fuzz.partial_ratio(str1, str2)
        elif 'Token Sort' in self.algorithm:
            return fuzz.token_sort_ratio(str1, str2)
        elif 'Token Set' in self.algorithm:
            return fuzz.token_set_ratio(str1, str2)
        else:  # Default: Ratio
            return fuzz.ratio(str1, str2)
    
    def find_best_match(self, 
                       target: str, 
                       candidates: pd.DataFrame, 
                       column: str,
                       secondary_column: Optional[str] = None,
                       weights: Tuple[float, float] = (1.0, 0.0)) -> Tuple[Optional[pd.Series], int, str]:
        """
        Find the best matching row from candidates.
        
        Args:
            target: Target string to match
            candidates: DataFrame with candidate rows
            column: Primary column to match against
            secondary_column: Optional secondary column for combined matching
            weights: Weight for primary and secondary columns
            
        Returns:
            Tuple of (best_match_row, score, match_type)
        """
        if candidates.empty:
            return None, 0, "No candidates"
        
        best_match = None
        best_score = 0
        match_type = "none"
        
        target_prepared = self.prepare_string(target)
        
        for idx, row in candidates.iterrows():
            if pd.isna(row.get(column)):
                continue
            
            # Check for exact match first
            candidate_prepared = self.prepare_string(row[column])
            if target_prepared == candidate_prepared:
                return row, 100, "Exact"
            
            # Fuzzy matching if enabled
            if self.enabled:
                primary_score = self.calculate_score(target, row[column])
                
                # Calculate combined score if secondary column provided
                if secondary_column and secondary_column in row and pd.notna(row.get(secondary_column)):
                    secondary_score = self.calculate_score(target, row[secondary_column])
                    score = (primary_score * weights[0]) + (secondary_score * weights[1])
                else:
                    score = primary_score
                
                if score > best_score and score >= self.threshold:
                    best_match = row
                    best_score = score
                    algorithm_name = self.algorithm.split(' ')[0]
                    match_type = f"Fuzzy-{algorithm_name} ({score:.0f}%)"
        
        return best_match, best_score, match_type
    
    def match_dataframes(self,
                        source_df: pd.DataFrame,
                        target_df: pd.DataFrame,
                        match_columns: Dict[str, str],
                        os_weight: float = 0.7,
                        hazard_weight: float = 0.3) -> pd.DataFrame:
        """
        Match rows between two dataframes using fuzzy matching.
        
        Args:
            source_df: Source dataframe
            target_df: Target dataframe to match against
            match_columns: Dictionary mapping source to target columns
            os_weight: Weight for operating scenario matching
            hazard_weight: Weight for hazard matching
            
        Returns:
            Matched dataframe with match scores and types
        """
        matched_results = []
        
        for _, source_row in source_df.iterrows():
            # Get the primary match column (operating scenario)
            os_col_source = match_columns.get('operating_scenario_source')
            os_col_target = match_columns.get('operating_scenario_target')
            
            if not os_col_source or pd.isna(source_row.get(os_col_source)):
                continue
            
            # Find best match
            best_match, score, match_type = self.find_best_match(
                source_row[os_col_source],
                target_df,
                os_col_target,
                match_columns.get('hazard_target'),
                (os_weight, hazard_weight)
            )
            
            # Create matched row
            matched_row = source_row.to_dict()
            matched_row['Match_Score'] = f"{score:.0f}%"
            matched_row['Match_Type'] = match_type
            
            # Add matched data if found
            if best_match is not None:
                for key, col in match_columns.items():
                    if key.endswith('_target') and col in best_match.index:
                        new_key = key.replace('_target', '_matched')
                        matched_row[new_key] = best_match[col]
            
            matched_results.append(matched_row)
        
        return pd.DataFrame(matched_results)