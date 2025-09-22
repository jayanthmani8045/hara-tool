"""
Excel processing module for HARA Tool
"""

import pandas as pd
from PySide6.QtCore import QThread, Signal
from typing import Optional, Dict, List, Tuple
from .matcher import FuzzyMatcher
from .constants import ASIL_DETERMINATION_TABLE, COLUMN_MAPPINGS, PARAMETER_RANGES


class ExcelProcessor(QThread):
    """Background thread for processing Excel files"""
    
    progress = Signal(int)
    log = Signal(str)
    finished = Signal(pd.DataFrame)
    error = Signal(str)
    
    def __init__(self, 
                 user_file: str,
                 os_sheet: str,
                 ra_sheet: str,
                 fuzzy_settings: Dict):
        """
        Initialize the Excel processor.
        
        Args:
            user_file: Path to the Excel file
            os_sheet: Operating Scenario sheet name
            ra_sheet: Risk Assessment sheet name
            fuzzy_settings: Dictionary with fuzzy matching settings
        """
        super().__init__()
        self.user_file = user_file
        self.os_sheet = os_sheet
        self.ra_sheet = ra_sheet
        self.result_df = None
        
        # Initialize fuzzy matcher with settings
        self.matcher = FuzzyMatcher(
            enabled=fuzzy_settings.get('enabled', True),
            threshold=fuzzy_settings.get('threshold', 80),
            algorithm=fuzzy_settings.get('algorithm', 'Ratio (Default)'),
            case_sensitive=fuzzy_settings.get('case_sensitive', False),
            strip_whitespace=fuzzy_settings.get('strip_whitespace', True)
        )
        
        self.os_weight = fuzzy_settings.get('os_weight', 70) / 100.0
        self.hazard_weight = fuzzy_settings.get('hazard_weight', 30) / 100.0
    
    def run(self):
        """Main processing logic"""
        try:
            # Step 1: Load Excel sheets
            self.log.emit(f"Loading Operating Scenario sheet: {self.os_sheet}")
            self.progress.emit(20)
            os_df = pd.read_excel(self.user_file, sheet_name=self.os_sheet)
            
            self.log.emit(f"Loading Risk Assessment sheet: {self.ra_sheet}")
            self.progress.emit(30)
            ra_df = pd.read_excel(self.user_file, sheet_name=self.ra_sheet)
            
            # Step 2: Process Operating Scenarios
            self.log.emit("Processing Operating Scenarios...")
            self.progress.emit(40)
            processed_df = self.process_operating_scenarios(os_df)
            
            # Step 3: Match with Risk Assessment
            self.log.emit(f"Matching scenarios (Fuzzy: {self.matcher.enabled}, Threshold: {self.matcher.threshold}%)")
            self.progress.emit(60)
            matched_df = self.match_risk_assessment(processed_df, ra_df)
            
            # Step 4: Determine ASIL values
            self.log.emit("Determining ASIL values...")
            self.progress.emit(80)
            final_df = self.determine_asil(matched_df)
            
            self.log.emit("Processing completed successfully")
            self.progress.emit(100)
            self.finished.emit(final_df)
            
        except Exception as e:
            self.error.emit(f"Error during processing: {str(e)}")
    
    def find_column_case_insensitive(self, df: pd.DataFrame, column_name: str) -> Optional[str]:
        """
        Find column name case-insensitively.
        
        Args:
            df: DataFrame to search
            column_name: Column name to find
            
        Returns:
            Column name if found, None otherwise
        """
        for col in df.columns:
            if str(col).strip().lower() == column_name.lower():
                return col
        return None
    
    def process_operating_scenarios(self, os_df: pd.DataFrame) -> pd.DataFrame:
        """
        Extract and process Operating Scenario data.
        
        Args:
            os_df: Operating Scenario dataframe
            
        Returns:
            Processed dataframe with scenarios
        """
        results = []
        
        # Find columns case-insensitively
        os_col = self.find_column_case_insensitive(os_df, 'operating scenario')
        e_col = self.find_column_case_insensitive(os_df, 'e')
        
        # Find hazard columns (multiple possible)
        hazard_cols = [col for col in os_df.columns 
                      if 'hazard' in str(col).lower() and col != os_col]
        
        if not os_col:
            raise ValueError("Operating Scenario column not found")
        
        self.log.emit(f"Found columns: Operating Scenario, E, {len(hazard_cols)} hazard columns")
        
        # Process each row
        for idx, row in os_df.iterrows():
            # Skip if Operating Scenario is empty
            if pd.isna(row.get(os_col, None)) or str(row.get(os_col, '')).strip() == '':
                continue
            
            operating_scenario = str(row[os_col]).strip()
            
            # Get E value (default to 4 if empty)
            e_value = 4
            if e_col and not pd.isna(row.get(e_col, None)):
                try:
                    e_value = int(row[e_col])
                except:
                    e_value = 4
            
            # Get hazards from all hazard columns
            hazards = []
            for h_col in hazard_cols:
                if not pd.isna(row.get(h_col, None)) and str(row.get(h_col, '')).strip():
                    hazards.append(str(row[h_col]).strip())
            
            # If no hazard found, we'll match from Risk Assessment later
            hazard = hazards[0] if hazards else None
            
            results.append({
                'Operating Scenario': operating_scenario,
                'E': e_value,
                'Hazard': hazard,
                'Original_Index': idx
            })
        
        self.log.emit(f"Processed {len(results)} operating scenarios")
        return pd.DataFrame(results)
    
    def match_risk_assessment(self, processed_df: pd.DataFrame, ra_df: pd.DataFrame) -> pd.DataFrame:
        """
        Match processed scenarios with Risk Assessment data.
        
        Args:
            processed_df: Processed operating scenarios
            ra_df: Risk Assessment dataframe
            
        Returns:
            Matched dataframe with risk assessment data
        """
        # Find columns in Risk Assessment
        ra_os_col = self.find_column_case_insensitive(ra_df, 'operating scenario')
        ra_hazard_col = self.find_column_case_insensitive(ra_df, 'hazard')
        ra_he_col = self.find_column_case_insensitive(ra_df, 'hazardous event')
        ra_details_col = self.find_column_case_insensitive(ra_df, 'details of hazardous event')
        ra_people_col = self.find_column_case_insensitive(ra_df, 'people at risk')
        ra_delta_col = self.find_column_case_insensitive(ra_df, 'δv') or self.find_column_case_insensitive(ra_df, 'deltav') or self.find_column_case_insensitive(ra_df, 'Δv')
        ra_s_col = self.find_column_case_insensitive(ra_df, 's')
        ra_s_rational_col = self.find_column_case_insensitive(ra_df, 'severity rational')
        ra_c_col = self.find_column_case_insensitive(ra_df, 'c')
        ra_c_rational_col = self.find_column_case_insensitive(ra_df, 'controllability rational')
        
        matched_results = []
        exact_matches = 0
        fuzzy_matches = 0
        no_matches = 0
        
        for _, row in processed_df.iterrows():
            os_text = row['Operating Scenario']
            hazard = row['Hazard']
            e_value = row['E']
            
            # Try exact match first
            best_match = None
            best_score = 0
            match_type = "none"
            match_details = ""
            
            # Prepare text for comparison
            os_text_prepared = self.matcher.prepare_string(os_text)
            
            # Exact matching phase
            for _, ra_row in ra_df.iterrows():
                if pd.isna(ra_row.get(ra_os_col, None)):
                    continue
                    
                ra_os_text = self.matcher.prepare_string(ra_row[ra_os_col])
                
                # Exact match
                if os_text_prepared == ra_os_text:
                    # If hazard specified, check it matches
                    if hazard and ra_hazard_col:
                        hazard_prepared = self.matcher.prepare_string(hazard)
                        if pd.notna(ra_row.get(ra_hazard_col)):
                            ra_hazard = self.matcher.prepare_string(ra_row[ra_hazard_col])
                            if hazard_prepared == ra_hazard:
                                best_match = ra_row
                                best_score = 100
                                match_type = "exact"
                                match_details = "Exact (OS+Hazard)"
                                break
                    else:
                        best_match = ra_row
                        best_score = 100
                        match_type = "exact"
                        match_details = "Exact (OS)"
                        break
            
            # Fuzzy matching phase (if enabled and no exact match)
            if best_score < 100 and self.matcher.enabled:
                for _, ra_row in ra_df.iterrows():
                    if pd.isna(ra_row.get(ra_os_col, None)):
                        continue
                    
                    ra_os_text = str(ra_row[ra_os_col]).strip()
                    
                    # Calculate fuzzy match score using selected algorithm
                    os_score = self.matcher.calculate_score(os_text, ra_os_text)
                    
                    # Calculate combined score with hazard if available
                    if hazard and ra_hazard_col and pd.notna(ra_row.get(ra_hazard_col)):
                        hazard_score = self.matcher.calculate_score(hazard, str(ra_row[ra_hazard_col]))
                        # Weighted average based on user settings
                        score = (os_score * self.os_weight) + (hazard_score * self.hazard_weight)
                    else:
                        score = os_score
                    
                    # Check if score meets threshold
                    if score > best_score and score >= self.matcher.threshold:
                        best_match = ra_row
                        best_score = score
                        match_type = "fuzzy"
                        algorithm_name = self.matcher.algorithm.split(' ')[0]
                        match_details = f"Fuzzy-{algorithm_name} ({score:.0f}%)"
            
            # Count match types
            if match_type == "exact":
                exact_matches += 1
            elif match_type == "fuzzy":
                fuzzy_matches += 1
            else:
                no_matches += 1
            
            # Prepare matched row with match information
            matched_row = {
                'Operating Scenario': os_text,
                'E': e_value,
                'Hazard': hazard if hazard else (best_match[ra_hazard_col] if best_match is not None and ra_hazard_col else ''),
                'Match_Type': match_details if match_details else 'No match',
                'Match_Score': f"{best_score:.0f}%" if best_score > 0 else "0%"
            }
            
            # Add fields from Risk Assessment if match found
            if best_match is not None:
                if ra_he_col:
                    matched_row['Hazardous Event'] = best_match.get(ra_he_col, '')
                if ra_details_col:
                    matched_row['Details of Hazardous event'] = best_match.get(ra_details_col, '')
                if ra_people_col:
                    matched_row['people at risk'] = best_match.get(ra_people_col, '')
                if ra_delta_col:
                    matched_row['Δv'] = best_match.get(ra_delta_col, '')
                
                # Only set S if valid value found (no default)
                if ra_s_col and pd.notna(best_match.get(ra_s_col)):
                    try:
                        s_val = int(best_match.get(ra_s_col))
                        if 0 <= s_val <= 3:  # Validate S range
                            matched_row['S'] = s_val
                        else:
                            matched_row['S'] = ''
                    except:
                        matched_row['S'] = ''
                else:
                    matched_row['S'] = ''
                
                if ra_s_rational_col:
                    matched_row['Severity Rational'] = best_match.get(ra_s_rational_col, '')
                
                # Only set C if valid value found (no default)
                if ra_c_col and pd.notna(best_match.get(ra_c_col)):
                    try:
                        c_val = int(best_match.get(ra_c_col))
                        if 0 <= c_val <= 3:  # Validate C range
                            matched_row['C'] = c_val
                        else:
                            matched_row['C'] = ''
                    except:
                        matched_row['C'] = ''
                else:
                    matched_row['C'] = ''
                    
                if ra_c_rational_col:
                    matched_row['Controllability Rational'] = best_match.get(ra_c_rational_col, '')
            else:
                # No match found - leave S and C empty (no defaults)
                matched_row.update({
                    'Hazardous Event': '',
                    'Details of Hazardous event': '',
                    'people at risk': '',
                    'Δv': '',
                    'S': '',
                    'Severity Rational': '',
                    'C': '',
                    'Controllability Rational': ''
                })
            
            matched_results.append(matched_row)
        
        algorithm_info = f" ({self.matcher.algorithm.split(' ')[0]})" if self.matcher.enabled else ""
        self.log.emit(f"Matching complete: {exact_matches} exact, {fuzzy_matches} fuzzy{algorithm_info}, {no_matches} unmatched")
        return pd.DataFrame(matched_results)
    
    def determine_asil(self, matched_df: pd.DataFrame) -> pd.DataFrame:
        """
        Determine ASIL values based on E, S, C parameters.
        
        Args:
            matched_df: Dataframe with matched scenarios
            
        Returns:
            Dataframe with ASIL values added
        """
        def get_asil_value(e: int, s: any, c: any) -> str:
            """Calculate ASIL value from E, S, C"""
            # Only calculate ASIL if S and C are valid values
            if pd.isna(s) or s == '' or pd.isna(c) or c == '':
                return ''  # Return empty if S or C is missing
            
            try:
                # Convert values to valid ranges
                e_val = min(max(int(e), 1), 4)
                s_val = int(s)
                c_val = int(c)
                
                # Validate ranges
                if not (0 <= s_val <= 3) or not (0 <= c_val <= 3):
                    return ''
                
                # Create lookup keys
                s_key = f"S{s_val}"
                e_key = f"E{e_val}"
                c_key = f"C{c_val}"
                
                # Lookup in embedded table
                if (s_key, e_key) in ASIL_DETERMINATION_TABLE:
                    return ASIL_DETERMINATION_TABLE[(s_key, e_key)].get(c_key, '')
            except:
                return ''
            
            return ''  # Return empty if lookup fails
        
        # Add ASIL column
        matched_df['ASIL'] = matched_df.apply(
            lambda row: get_asil_value(row['E'], row['S'], row['C']),
            axis=1
        )
        
        # Count ASIL distribution
        asil_values = matched_df['ASIL'][matched_df['ASIL'] != '']
        if len(asil_values) > 0:
            asil_counts = asil_values.value_counts()
            self.log.emit(f"ASIL determination: {', '.join([f'{val}={count}' for val, count in asil_counts.items()])}")
        else:
            self.log.emit("ASIL determination: No valid S/C values found for ASIL calculation")
        
        return matched_df
    
    def find_column_case_insensitive(self, df: pd.DataFrame, column_name: str) -> Optional[str]:
        """
        Find column name case-insensitively.
        
        Args:
            df: DataFrame to search
            column_name: Column name to find
            
        Returns:
            Column name if found, None otherwise
        """
        for col in df.columns:
            if str(col).strip().lower() == column_name.lower():
                return col
        return None
    
    def diagnose_s_c_values(self, df: pd.DataFrame, s_col: str, c_col: str) -> Dict[str, any]:
        """
        Diagnose S and C values in dataframe.
        
        Args:
            df: DataFrame to analyze
            s_col: S column name
            c_col: C column name
            
        Returns:
            Dictionary with diagnostic information
        """
        diag = {
            's_found': s_col is not None,
            'c_found': c_col is not None,
            's_valid_count': 0,
            'c_valid_count': 0,
            's_invalid_values': [],
            'c_invalid_values': [],
            'total_rows': len(df)
        }
        
        if s_col and s_col in df.columns:
            s_values = df[s_col]
            for val in s_values:
                if pd.notna(val):
                    try:
                        s_int = int(val)
                        if 0 <= s_int <= 3:
                            diag['s_valid_count'] += 1
                        else:
                            diag['s_invalid_values'].append(str(val))
                    except:
                        diag['s_invalid_values'].append(str(val))
        
        if c_col and c_col in df.columns:
            c_values = df[c_col]
            for val in c_values:
                if pd.notna(val):
                    try:
                        c_int = int(val)
                        if 0 <= c_int <= 3:
                            diag['c_valid_count'] += 1
                        else:
                            diag['c_invalid_values'].append(str(val))
                    except:
                        diag['c_invalid_values'].append(str(val))
        
        return diag
    
    def process_operating_scenarios(self, os_df: pd.DataFrame) -> pd.DataFrame:
        """
        Extract and process Operating Scenario data.
        
        Args:
            os_df: Operating Scenario dataframe
            
        Returns:
            Processed dataframe with scenarios
        """
        results = []
        
        # Find columns case-insensitively
        os_col = self.find_column_case_insensitive(os_df, 'operating scenario')
        e_col = self.find_column_case_insensitive(os_df, 'e')
        
        # Find hazard columns (multiple possible)
        hazard_cols = [col for col in os_df.columns 
                      if 'hazard' in str(col).lower() and col != os_col]
        
        if not os_col:
            raise ValueError("Operating Scenario column not found")
        
        self.log.emit(f"Found columns: Operating Scenario, E, {len(hazard_cols)} hazard columns")
        
        # Process each row
        for idx, row in os_df.iterrows():
            # Skip if Operating Scenario is empty
            if pd.isna(row.get(os_col, None)) or str(row.get(os_col, '')).strip() == '':
                continue
            
            operating_scenario = str(row[os_col]).strip()
            
            # Get E value (default to 4 if empty)
            e_value = 4
            if e_col and not pd.isna(row.get(e_col, None)):
                try:
                    e_value = int(row[e_col])
                except:
                    e_value = 4
            
            # Get hazards from all hazard columns
            hazards = []
            for h_col in hazard_cols:
                if not pd.isna(row.get(h_col, None)) and str(row.get(h_col, '')).strip():
                    hazards.append(str(row[h_col]).strip())
            
            # If no hazard found, we'll match from Risk Assessment later
            hazard = hazards[0] if hazards else None
            
            results.append({
                'Operating Scenario': operating_scenario,
                'E': e_value,
                'Hazard': hazard,
                'Original_Index': idx
            })
        
        self.log.emit(f"Processed {len(results)} operating scenarios")
        return pd.DataFrame(results)
    
    def match_risk_assessment(self, processed_df: pd.DataFrame, ra_df: pd.DataFrame) -> pd.DataFrame:
        """
        Match processed scenarios with Risk Assessment data.
        
        Args:
            processed_df: Processed operating scenarios
            ra_df: Risk Assessment dataframe
            
        Returns:
            Matched dataframe with risk assessment data
        """
        # Find columns in Risk Assessment
        ra_os_col = self.find_column_case_insensitive(ra_df, 'operating scenario')
        ra_hazard_col = self.find_column_case_insensitive(ra_df, 'hazard')
        ra_he_col = self.find_column_case_insensitive(ra_df, 'hazardous event')
        ra_details_col = self.find_column_case_insensitive(ra_df, 'details of hazardous event')
        ra_people_col = self.find_column_case_insensitive(ra_df, 'people at risk')
        ra_delta_col = self.find_column_case_insensitive(ra_df, 'δv') or self.find_column_case_insensitive(ra_df, 'deltav') or self.find_column_case_insensitive(ra_df, 'Δv')
        ra_s_col = self.find_column_case_insensitive(ra_df, 's')
        ra_s_rational_col = self.find_column_case_insensitive(ra_df, 'severity rational')
        ra_c_col = self.find_column_case_insensitive(ra_df, 'c')
        ra_c_rational_col = self.find_column_case_insensitive(ra_df, 'controllability rational')
        
        matched_results = []
        exact_matches = 0
        fuzzy_matches = 0
        no_matches = 0
        
        for _, row in processed_df.iterrows():
            os_text = row['Operating Scenario']
            hazard = row['Hazard']
            e_value = row['E']
            
            # Try exact match first
            best_match = None
            best_score = 0
            match_type = "none"
            match_details = ""
            
            # Prepare text for comparison
            os_text_prepared = self.matcher.prepare_string(os_text)
            
            # Exact matching phase
            for _, ra_row in ra_df.iterrows():
                if pd.isna(ra_row.get(ra_os_col, None)):
                    continue
                    
                ra_os_text = self.matcher.prepare_string(ra_row[ra_os_col])
                
                # Exact match
                if os_text_prepared == ra_os_text:
                    # If hazard specified, check it matches
                    if hazard and ra_hazard_col:
                        hazard_prepared = self.matcher.prepare_string(hazard)
                        if pd.notna(ra_row.get(ra_hazard_col)):
                            ra_hazard = self.matcher.prepare_string(ra_row[ra_hazard_col])
                            if hazard_prepared == ra_hazard:
                                best_match = ra_row
                                best_score = 100
                                match_type = "exact"
                                match_details = "Exact (OS+Hazard)"
                                break
                    else:
                        best_match = ra_row
                        best_score = 100
                        match_type = "exact"
                        match_details = "Exact (OS)"
                        break
            
            # Fuzzy matching phase (if enabled and no exact match)
            if best_score < 100 and self.matcher.enabled:
                for _, ra_row in ra_df.iterrows():
                    if pd.isna(ra_row.get(ra_os_col, None)):
                        continue
                    
                    ra_os_text = str(ra_row[ra_os_col]).strip()
                    
                    # Calculate fuzzy match score using selected algorithm
                    os_score = self.matcher.calculate_score(os_text, ra_os_text)
                    
                    # Calculate combined score with hazard if available
                    if hazard and ra_hazard_col and pd.notna(ra_row.get(ra_hazard_col)):
                        hazard_score = self.matcher.calculate_score(hazard, str(ra_row[ra_hazard_col]))
                        # Weighted average based on user settings
                        score = (os_score * self.os_weight) + (hazard_score * self.hazard_weight)
                    else:
                        score = os_score
                    
                    # Check if score meets threshold
                    if score > best_score and score >= self.matcher.threshold:
                        best_match = ra_row
                        best_score = score
                        match_type = "fuzzy"
                        algorithm_name = self.matcher.algorithm.split(' ')[0]
                        match_details = f"Fuzzy-{algorithm_name} ({score:.0f}%)"
            
            # Count match types
            if match_type == "exact":
                exact_matches += 1
            elif match_type == "fuzzy":
                fuzzy_matches += 1
            else:
                no_matches += 1
            
            # Prepare matched row with match information
            matched_row = {
                'Operating Scenario': os_text,
                'E': e_value,
                'Hazard': hazard if hazard else (best_match[ra_hazard_col] if best_match is not None and ra_hazard_col else ''),
                'Match_Type': match_details if match_details else 'No match',
                'Match_Score': f"{best_score:.0f}%" if best_score > 0 else "0%"
            }
            
            # Add fields from Risk Assessment if match found
            if best_match is not None:
                if ra_he_col:
                    matched_row['Hazardous Event'] = best_match.get(ra_he_col, '')
                if ra_details_col:
                    matched_row['Details of Hazardous event'] = best_match.get(ra_details_col, '')
                if ra_people_col:
                    matched_row['people at risk'] = best_match.get(ra_people_col, '')
                if ra_delta_col:
                    matched_row['Δv'] = best_match.get(ra_delta_col, '')
                
                # Only set S if valid value found (no default)
                if ra_s_col and pd.notna(best_match.get(ra_s_col)):
                    try:
                        s_val = int(best_match.get(ra_s_col))
                        if 0 <= s_val <= 3:  # Validate S range
                            matched_row['S'] = s_val
                        else:
                            matched_row['S'] = ''
                    except:
                        matched_row['S'] = ''
                else:
                    matched_row['S'] = ''
                
                if ra_s_rational_col:
                    matched_row['Severity Rational'] = best_match.get(ra_s_rational_col, '')
                
                # Only set C if valid value found (no default)
                if ra_c_col and pd.notna(best_match.get(ra_c_col)):
                    try:
                        c_val = int(best_match.get(ra_c_col))
                        if 0 <= c_val <= 3:  # Validate C range
                            matched_row['C'] = c_val
                        else:
                            matched_row['C'] = ''
                    except:
                        matched_row['C'] = ''
                else:
                    matched_row['C'] = ''
                    
                if ra_c_rational_col:
                    matched_row['Controllability Rational'] = best_match.get(ra_c_rational_col, '')
            else:
                # No match found - leave S and C empty (no defaults)
                matched_row.update({
                    'Hazardous Event': '',
                    'Details of Hazardous event': '',
                    'people at risk': '',
                    'Δv': '',
                    'S': '',
                    'Severity Rational': '',
                    'C': '',
                    'Controllability Rational': ''
                })
            
            matched_results.append(matched_row)
        
        algorithm_info = f" ({self.matcher.algorithm.split(' ')[0]})" if self.matcher.enabled else ""
        self.log.emit(f"Matching complete: {exact_matches} exact, {fuzzy_matches} fuzzy{algorithm_info}, {no_matches} unmatched")
        return pd.DataFrame(matched_results)
    
    def determine_asil(self, matched_df: pd.DataFrame) -> pd.DataFrame:
        """
        Determine ASIL values based on E, S, C parameters.
        
        Args:
            matched_df: Dataframe with matched scenarios
            
        Returns:
            Dataframe with ASIL values added
        """
        def get_asil_value(e: int, s: any, c: any) -> str:
            """Calculate ASIL value from E, S, C"""
            # Check if S and C are valid
            if pd.isna(s) or s == '' or pd.isna(c) or c == '':
                return ''
            
            try:
                # Convert and validate values
                e_val = max(1, min(4, int(e)))
                s_val = int(s)
                c_val = int(c)
                
                # Validate ranges
                if not (0 <= s_val <= 3) or not (0 <= c_val <= 3):
                    return ''
                
                # Create lookup keys
                s_key = f"S{s_val}"
                e_key = f"E{e_val}"
                c_key = f"C{c_val}"
                
                # Lookup in table
                if (s_key, e_key) in ASIL_DETERMINATION_TABLE:
                    return ASIL_DETERMINATION_TABLE[(s_key, e_key)].get(c_key, '')
                
            except Exception:
                return ''
            
            return ''
        
        # Calculate ASIL for each row
        matched_df['ASIL'] = matched_df.apply(
            lambda row: get_asil_value(row['E'], row['S'], row['C']),
            axis=1
        )
        
        # Log ASIL distribution
        asil_values = matched_df['ASIL'][matched_df['ASIL'] != '']
        if len(asil_values) > 0:
            asil_counts = asil_values.value_counts()
            self.log.emit(f"ASIL determination: {', '.join([f'{val}={count}' for val, count in asil_counts.items()])}")
        else:
            self.log.emit("ASIL determination: No valid S/C values found for ASIL calculation")
        
        return matched_df