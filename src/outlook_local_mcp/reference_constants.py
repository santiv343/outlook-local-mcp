"""Short reference wire names and entropy; resource bounds reuse cursor limits."""

from .enums import EReferenceKind

REFERENCE_PREFIX = "olm_"
REFERENCE_RANDOM_BYTES = 9
REFERENCE_FIELDS = {
    "store_id": EReferenceKind.STORE,
    "folder_id": EReferenceKind.FOLDER,
    "parent_folder_id": EReferenceKind.FOLDER,
    "entry_id": EReferenceKind.ITEM,
}
