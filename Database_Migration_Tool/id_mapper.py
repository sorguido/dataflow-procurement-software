"""
DataFlow Database Migration Tool - ID Mapper
Handles ID remapping from v1.1.0 sequential IDs to v2.0.0 year-based IDs
"""

from datetime import datetime


class IDMapper:
    """
    Maps old v1.1.0 IDs to new v2.0.0 year-based IDs.
    
    v2.0.0 uses format: YYXXXXX where YY = year, XXXXX = sequence
    Example: 2500001 for first RfQ in 2025
    """
    
    def __init__(self, target_year=None):
        """
        Initialize ID mapper.
        
        Args:
            target_year: Year to use for new IDs (default: current year)
        """
        self.target_year = target_year or datetime.now().year
        self.id_map = {}  # old_id -> new_id
        self.next_sequence = 1
        self.base_id = (self.target_year % 100) * 100000  # e.g., 2500000 for 2025
    
    def generate_new_id(self, old_id):
        """
        Generate a new year-based ID for an old v1.1.0 ID.
        
        Args:
            old_id: Original v1.1.0 ID (any integer)
            
        Returns:
            int: New v2.0.0 ID in format YYXXXXX
        """
        if old_id in self.id_map:
            return self.id_map[old_id]
        
        new_id = self.base_id + self.next_sequence
        self.id_map[old_id] = new_id
        self.next_sequence += 1
        
        return new_id
    
    def get_mapping(self, old_id):
        """
        Get the new ID for an old ID (must have been generated first).
        
        Args:
            old_id: Original v1.1.0 ID
            
        Returns:
            int: Mapped v2.0.0 ID or None if not mapped yet
        """
        return self.id_map.get(old_id)
    
    def get_all_mappings(self):
        """
        Get all ID mappings.
        
        Returns:
            dict: Complete mapping {old_id: new_id}
        """
        return self.id_map.copy()
    
    def total_mapped(self):
        """
        Get count of mapped IDs.
        
        Returns:
            int: Number of IDs mapped
        """
        return len(self.id_map)
