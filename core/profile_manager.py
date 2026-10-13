import os
import json
import logging
from typing import Dict, List, Any, Tuple, Optional
import pandas as pd

from config import DEFAULT_VARIABLE_VALUES, ALL_SMAP_VARIABLES, IMAGE_SLOTS

logger = logging.getLogger("profile_manager")

class ProfileManager:
    """Manages profile loading, saving, JSON presets, Excel import/export,
    and default sample profile generation."""

    @staticmethod
    def get_sample_smap_profile() -> Dict[str, Any]:
        """Returns default SMAP corporate data dictionary with all 78 variables populated."""
        profile = {
            "variables": dict(DEFAULT_VARIABLE_VALUES),
            "images": {}
        }
        for tag_name, slot in IMAGE_SLOTS.items():
            profile["images"][tag_name] = {
                "path": "",
                "width_mm": slot["default_width_mm"],
                "height_mm": slot["default_height_mm"]
            }
        return profile

    @staticmethod
    def save_profile_json(data: Dict[str, Any], json_path: str) -> Tuple[bool, str]:
        """Saves form state (variables & image paths/dimensions) to JSON preset file."""
        try:
            os.makedirs(os.path.dirname(os.path.abspath(json_path)), exist_ok=True)
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            return True, json_path
        except Exception as e:
            logger.error(f"Failed to save JSON profile '{json_path}': {e}")
            return False, str(e)

    @staticmethod
    def load_profile_json(json_path: str) -> Tuple[bool, Optional[Dict[str, Any]], str]:
        """Loads form state from JSON preset file."""
        if not os.path.exists(json_path):
            return False, None, f"JSON profile file not found: {json_path}"

        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            # Ensure all variables are present with fallbacks
            loaded_vars = data.get("variables", {})
            merged_vars = dict(DEFAULT_VARIABLE_VALUES)
            merged_vars.update(loaded_vars)

            # Ensure both jenis_usaha and jenis_usaha_bu are synchronized
            if "jenis_usaha_bu" in loaded_vars and ("jenis_usaha" not in loaded_vars or not loaded_vars["jenis_usaha"]):
                merged_vars["jenis_usaha"] = loaded_vars["jenis_usaha_bu"]
            if "jenis_usaha" in loaded_vars and ("jenis_usaha_bu" not in loaded_vars or not loaded_vars["jenis_usaha_bu"]):
                merged_vars["jenis_usaha_bu"] = loaded_vars["jenis_usaha"]

            loaded_images = data.get("images", {})
            merged_images = {}
            for tag_name, slot in IMAGE_SLOTS.items():
                slot_info = loaded_images.get(tag_name, {})
                merged_images[tag_name] = {
                    "path": slot_info.get("path", ""),
                    "width_mm": slot_info.get("width_mm", slot["default_width_mm"]),
                    "height_mm": slot_info.get("height_mm", slot["default_height_mm"])
                }

            result_profile = {
                "variables": merged_vars,
                "images": merged_images
            }
            return True, result_profile, json_path

        except Exception as e:
            logger.error(f"Failed to load JSON profile '{json_path}': {e}")
            return False, None, str(e)

    @staticmethod
    def export_profile_excel(
        data_list: List[Dict[str, Any]],
        excel_path: str
    ) -> Tuple[bool, str]:
        """Exports single or multi-row list of variable dictionaries to an Excel (.xlsx) file.
        
        Args:
            data_list: List of dictionaries containing variable key-value pairs.
            excel_path: Destination .xlsx file path.
        """
        try:
            os.makedirs(os.path.dirname(os.path.abspath(excel_path)), exist_ok=True)
            df = pd.DataFrame(data_list)
            
            # Reorder columns to match canonical variables and group dynamic suffix fields
            import re
            ordered_cols = []
            extra_cols = [c for c in df.columns if c not in ALL_SMAP_VARIABLES]
            remaining_extra = set(extra_cols)

            i = 0
            while i < len(ALL_SMAP_VARIABLES):
                base_col = ALL_SMAP_VARIABLES[i]
                
                match = re.search(r'(\d+)$', base_col)
                if match and int(match.group(1)) == 1:
                    prefix = base_col[:-len(match.group(1))]
                    group_prefixes = [prefix]
                    
                    j = i + 1
                    base_name_no_underscore = prefix.rstrip('_')
                    while j < len(ALL_SMAP_VARIABLES):
                        next_col = ALL_SMAP_VARIABLES[j]
                        next_match = re.search(r'(\d+)$', next_col)
                        if next_match and int(next_match.group(1)) == 1 and base_name_no_underscore in next_col:
                            group_prefixes.append(next_col[:-len(next_match.group(1))])
                            j += 1
                        else:
                            break
                            
                    max_idx = 1
                    for extra in remaining_extra:
                        for p in group_prefixes:
                            if extra.startswith(p):
                                suffix = extra[len(p):]
                                if suffix.isdigit():
                                    max_idx = max(max_idx, int(suffix))
                                    
                    for idx in range(1, max_idx + 1):
                        for p in group_prefixes:
                            col_name = f'{p}{idx}'
                            if col_name in df.columns:
                                if col_name not in ordered_cols:
                                    ordered_cols.append(col_name)
                                if col_name in remaining_extra:
                                    remaining_extra.remove(col_name)
                                    
                    i = j
                else:
                    if base_col in df.columns:
                        ordered_cols.append(base_col)
                        
                    prefix = base_col + '_'
                    related = [x for x in remaining_extra if x.startswith(prefix) and x[len(prefix):].isdigit()]
                    related.sort(key=lambda x: int(x[len(prefix):]))
                    for r in related:
                        ordered_cols.append(r)
                        remaining_extra.remove(r)
                        
                    i += 1

            # Any remaining extra columns go to the end
            ordered_cols.extend(sorted(list(remaining_extra)))
            
            df = df[ordered_cols]

            # Transpose dataframe so variables are rows, easier to edit
            df_t = df.transpose()
            df_t.reset_index(inplace=True)
            
            # Generate new column names: Variable, Data 1, Data 2, ...
            col_names = ["Variable"] + [f"Data {i+1}" for i in range(len(df_t.columns) - 1)]
            df_t.columns = col_names

            df_t.to_excel(excel_path, index=False, engine="openpyxl")
            return True, excel_path
        except Exception as e:
            logger.error(f"Failed to export Excel profile '{excel_path}': {e}")
            return False, str(e)

    @staticmethod
    def import_profile_excel(excel_path: str) -> Tuple[bool, List[Dict[str, str]], str]:
        """Imports multi-row batch data from Excel (.xlsx) file.
        
        Returns:
            Tuple[bool, List[Dict[str, str]], str]: (Success status, List of row dicts, Message)
        """
        if not os.path.exists(excel_path):
            return False, [], f"Excel file not found: {excel_path}"

        try:
            df = pd.read_excel(excel_path, engine="openpyxl", dtype=str)
            df = df.fillna("")  # Replace NaN with empty string
            
            # Detect if it's the transposed format (where first column is 'Variable')
            first_col_name = str(df.columns[0]).strip().lower()
            if first_col_name == "variable":
                df.set_index(df.columns[0], inplace=True)
                df = df.transpose()

            records = []
            for _, row in df.iterrows():
                # Skip empty columns/rows if the user left extra blank data columns
                if all(str(val).strip() == "" for val in row.values):
                    continue

                row_dict = dict(DEFAULT_VARIABLE_VALUES)  # Start with defaults
                for col in df.columns:
                    col_key = str(col).strip().lower()
                    if col_key in ALL_SMAP_VARIABLES:
                        val = str(row[col]).strip()
                        if val:
                            row_dict[col_key] = val
                    else:
                        row_dict[str(col).strip()] = str(row[col]).strip()
                if "jenis_usaha_bu" in row_dict and ("jenis_usaha" not in row_dict or not row_dict["jenis_usaha"]):
                    row_dict["jenis_usaha"] = row_dict["jenis_usaha_bu"]
                if "jenis_usaha" in row_dict and ("jenis_usaha_bu" not in row_dict or not row_dict["jenis_usaha_bu"]):
                    row_dict["jenis_usaha_bu"] = row_dict["jenis_usaha"]
                records.append(row_dict)

            return True, records, f"Successfully loaded {len(records)} row(s) from Excel."
        except Exception as e:
            logger.error(f"Failed to import Excel profile '{excel_path}': {e}")
            return False, [], str(e)
