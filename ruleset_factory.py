import importlib

class RulesetFactory:
    """
    Dynamically routes engine pipelines to ruleset-specific parsing scripts
    and standardizes asset file-naming conventions seamlessly.
    """
    def __init__(self, ruleset_name):
        self.ruleset_raw = ruleset_name if ruleset_name else "Unknown"
        # Normalize to lower ID: 'FrontierSpace' -> 'fs'
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
        """Standardizes file names (e.g., 'story' -> 'FS_Story_Template.docx')"""
        return f"{self.docx_prefix}{component_name.capitalize()}_Template.docx"

    def execute_parser(self, component_name, xml_root, *args, **kwargs):
        """
        Dynamically locates, imports, and calls the designated module parser function.
        Example: 'vehicle' -> calls fs_vehicle_parser.get_vehicle_catalog(xml_root)
        """
        module_name = f"{self.py_prefix}{component_name.lower()}_parser"
        function_name = f"get_{component_name.lower()}_catalog"
        
        try:
            module = importlib.import_module(module_name)
            parser_func = getattr(module, function_name)
            return parser_func(xml_root, *args, **kwargs)
        except ModuleNotFoundError:
            print(f"[INFO] No specialized rule-module parser found for '{module_name}.py'. Skipping pipeline stage.")
            return None
        except AttributeError:
            print(f"[!] Engine Error: Found script module '{module_name}', but function '{function_name}' is missing.")
            return None