import importlib
import os
import fg_parser  # Import your global parser fallback module cleanly

class RulesetFactory:
    """
    Dynamically routes engine pipelines to ruleset-specific parsing scripts
    and standardizes asset file-naming conventions seamlessly.
    """
    def __init__(self, ruleset_name):
        self.ruleset_raw = ruleset_name if ruleset_name else "Unknown"
        self.ruleset_id = self.ruleset_raw.strip().lower()
        
        if "frontierspace" in self.ruleset_id:
            self.py_prefix = "fs_"
            self.docx_prefix = "FS_"
        else:
            self.py_prefix = ""
            self.docx_prefix = ""

    def get_ruleset_name(self):
        return self.ruleset_raw

    def get_template_path(self, component_name):
        """Standardizes file names uniformly (e.g., 'starship' -> 'FS_Starship_Template.docx')"""
        return f"{self.docx_prefix}{component_name.capitalize()}_Template.docx"

    def execute_parser(self, component_name, xml_root, *args, **kwargs):
        """
        Dynamically locates, imports, and calls the designated module parser function.
        Safely routes legacy CoreRPG components to fg_parser if independent files are missing.
        """
        if component_name == "vehicle":
            module_name = "fs_vehicle_parser"
        elif component_name == "starship":
            module_name = "fs_starships_parser"
        elif component_name in ["npc", "item"]:
            module_name = "fg_parser"
        else:
            module_name = f"{component_name.lower()}_parser"
            
        try:
            module = importlib.import_module(module_name)
            
            # Check singular function signature (e.g., get_starship_catalog)
            function_name = f"get_{component_name.lower()}_catalog"
            parser_func = getattr(module, function_name, None)
            
            # Fallback to plural function signature (e.g., get_starships_catalog, get_tables_catalog)
            if parser_func is None:
                function_name = f"get_{component_name.lower()}s_catalog"
                parser_func = getattr(module, function_name)
                
            return parser_func(xml_root, *args, **kwargs)
        except (ModuleNotFoundError, AttributeError):
            try:
                function_name = f"get_{component_name.lower()}_catalog"
                parser_func = getattr(fg_parser, function_name, None)
                if parser_func is None:
                    function_name = f"get_{component_name.lower()}s_catalog"
                    parser_func = getattr(fg_parser, function_name)
                return parser_func(xml_root, *args, **kwargs)
            except AttributeError:
                print(f"[INFO] No specialized rule-module parser found for '{component_name}'. Skipping pipeline stage.")
                return None

    def get_renderer_module(self, component_name):
        """
        Dynamically imports and returns the appropriate renderer module.
        """
        if component_name == "starship":
            module_name = "fs_starships_renderer"
        else:
            module_name = f"{component_name.lower()}_renderer"

        try:
            return importlib.import_module(module_name)
        except ModuleNotFoundError:
            print(f"[!] Engine Error: Renderer module '{module_name}.py' could not be found.")
            return None