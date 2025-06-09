from .exLoadoutCheckpointLoader import exLoadoutCheckpointLoader
from .exLoadoutSelector import exLoadoutSelector
from .exLoadoutA import exLoadoutSeg
from .exLoadoutG import exLoadoutSeg2
from .exLoadoutReadColumn import exLoadoutReadColumn
from .exLoadoutEditCell import exLoadoutEditCell

NODE_CLASS_MAPPINGS = {
    "exCheckpointLoader": exLoadoutCheckpointLoader,
    "dropdowns": exLoadoutSelector,
    "exSeg": exLoadoutSeg,
    "exSeg2": exLoadoutSeg2,
    "exLoadoutReadColumn": exLoadoutReadColumn,
    "exLoadoutEditCell": exLoadoutEditCell,
}

NODE_DISPLAY_NAME_MAPPINGS = {
    "exCheckpointLoader": "exLoadoutCheckpointLoader",
    "dropdowns": "exLoadout Selector",
    "exSeg": "exLoadoutA",
    "exSeg2": "exLoadoutG",
    "exLoadoutReadColumn": "exLoadoutReadColumn",
    "exLoadoutEditCell": "exLoadoutEditCell",
}

print("ExcelPicker Node Loaded Successfully")
