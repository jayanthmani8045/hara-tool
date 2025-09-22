"""
Constants and configuration for HARA Automation Tool
"""

# ASIL Determination Table (ISO 26262)
ASIL_DETERMINATION_TABLE = {
    ('S1', 'E1'): {'C0': 'QM', 'C1': 'QM', 'C2': 'QM', 'C3': 'QM'},
    ('S1', 'E2'): {'C0': 'QM', 'C1': 'QM', 'C2': 'QM', 'C3': 'QM'},
    ('S1', 'E3'): {'C0': 'QM', 'C1': 'QM', 'C2': 'QM', 'C3': 'A'},
    ('S1', 'E4'): {'C0': 'QM', 'C1': 'QM', 'C2': 'A', 'C3': 'B'},
    ('S2', 'E1'): {'C0': 'QM', 'C1': 'QM', 'C2': 'QM', 'C3': 'QM'},
    ('S2', 'E2'): {'C0': 'QM', 'C1': 'QM', 'C2': 'QM', 'C3': 'A'},
    ('S2', 'E3'): {'C0': 'QM', 'C1': 'QM', 'C2': 'A', 'C3': 'B'},
    ('S2', 'E4'): {'C0': 'A', 'C1': 'A', 'C2': 'B', 'C3': 'C'},
    ('S3', 'E1'): {'C0': 'QM', 'C1': 'QM', 'C2': 'A', 'C3': 'B'},
    ('S3', 'E2'): {'C0': 'QM', 'C1': 'QM', 'C2': 'B', 'C3': 'C'},
    ('S3', 'E3'): {'C0': 'A', 'C1': 'A', 'C2': 'B', 'C3': 'C'},
    ('S3', 'E4'): {'C0': 'B', 'C1': 'B', 'C2': 'C', 'C3': 'D'},
}

# ASIL color mapping for UI display
ASIL_COLORS = {
    'QM': (46, 125, 50),       # Green
    'A': (255, 193, 7),        # Amber
    'B': (255, 152, 0),        # Orange
    'C': (255, 87, 34),        # Deep Orange
    'D': (244, 67, 54)         # Red
}

# Default settings
DEFAULT_SETTINGS = {
    'fuzzy_enabled': True,
    'fuzzy_threshold': 80,
    'fuzzy_algorithm': 'Ratio (Default)',
    'case_sensitive': False,
    'strip_whitespace': True,
    'os_weight': 70,
    'hazard_weight': 30
}

# Column mappings for Excel sheets
COLUMN_MAPPINGS = {
    'operating_scenario': ['operating scenario', 'os', 'operating_scenario', 'operating scenarios'],
    'hazard': ['hazard', 'hazards'],
    'hazardous_event': ['hazardous event', 'he', 'hazardous_event'],
    'details': ['details of hazardous event', 'details', 'he_details'],
    'people_at_risk': ['people at risk', 'people', 'persons_at_risk'],
    'delta_v': ['δv', 'deltav', 'Δv', 'delta_v'],
    'severity': ['s', 'severity', 'S', 'Severity', 'severity rating', 'severity level'],
    'exposure': ['e', 'exposure', 'E', 'Exposure', 'exposure rating', 'exposure level'],
    'controllability': ['c', 'controllability', 'C', 'Controllability', 'controllability rating', 'controllability level'],
    'severity_rationale': ['severity rational', 'severity_rationale', 's_rationale', 'severity rationale'],
    'controllability_rationale': ['controllability rational', 'controllability_rationale', 'c_rationale', 'controllability rationale']
}

# Valid parameter ranges
PARAMETER_RANGES = {
    'severity': (0, 3),
    'exposure': (1, 4),
    'controllability': (0, 3)
}