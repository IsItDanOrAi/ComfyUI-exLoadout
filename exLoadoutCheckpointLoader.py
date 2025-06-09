import openpyxl
import os
from folder_paths import get_filename_list, get_full_path_or_raise, get_folder_paths
import comfy.sd

class exLoadoutCheckpointLoader:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "excel_path": ("STRING", {"default": os.path.join("exLoadoutList.xlsx")} ),
                "sheet_name": ("STRING", {"default": "MODELS"}),
                "loadout_name": ("STRING", {"default": ""}),  # Dropdown list populated from Column A
                "clip_type": (["stable_diffusion", "stable_cascade", "sd3", "stable_audio", "mochi", "ltxv", "pixart", "cosmos", "lumina2", "wan"],),
            },
        }

    RETURN_TYPES = ("MODEL", "CLIP", "VAE", "STRING")
    RETURN_NAMES = ("model", "clip", "vae", "Output")
    FUNCTION = "exLoadoutCheckpointLoader"
    CATEGORY = "exLoadout"
    DESCRIPTION = ("Loads a checkpoint model by reading its name from Column B, "
                   "CLIP from Column C, and VAE from Column D in an Excel file. "
                   "Each row is identified by a 'Loadout' name from Column A.")

    def exLoadoutCheckpointLoader(self, excel_path, sheet_name, loadout_name, clip_type):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        full_excel_path = os.path.join(current_dir, excel_path)
        if not os.path.exists(full_excel_path):
            raise FileNotFoundError(f"Excel file not found: {full_excel_path}")
        
        workbook = openpyxl.load_workbook(full_excel_path)
        if sheet_name not in workbook.sheetnames:
            workbook.close()
            raise ValueError(f"Sheet '{sheet_name}' not found in the Excel file")
        
        sheet = workbook[sheet_name]
        found_row = None
        
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
            if row[0].value and str(row[0].value).strip() == loadout_name:
                found_row = row
                break

        workbook.close()
        
        if not found_row:
            raise ValueError(f"Loadout '{loadout_name}' not found in Column A.")
        
        # Load checkpoint model (Column B)
        if not found_row[1].value:
            raise ValueError(f"No valid checkpoint name found for Loadout '{loadout_name}' in Column B.")
        ckpt_name = str(found_row[1].value).strip()

        allowed_ckpts = get_filename_list("checkpoints")
        if ckpt_name not in allowed_ckpts:
            raise ValueError(f"Checkpoint '{ckpt_name}' is not in the allowed checkpoints list.")
        
        ckpt_path = get_full_path_or_raise("checkpoints", ckpt_name)
        out = comfy.sd.load_checkpoint_guess_config(
            ckpt_path,
            output_vae=True,
            output_clip=True,
            embedding_directory=get_folder_paths("embeddings")
        )
        model, clip, vae = out[:3]

        # Load CLIP (Column C)
        clip_name = "Default"
        if len(found_row) > 2 and found_row[2].value:
            temp_clip_name = str(found_row[2].value).strip()
            allowed_clip = get_filename_list("text_encoders")
            if temp_clip_name in allowed_clip:
                try:
                    clip_path = get_full_path_or_raise("text_encoders", temp_clip_name)
                    clip_loader = comfy.sd.load_clip
                    clip = clip_loader(
                        ckpt_paths=[clip_path],
                        embedding_directory=get_folder_paths("embeddings"),
                        clip_type=clip_type
                    )
                    clip_name = temp_clip_name
                except Exception as e:
                    print(f"Warning: Failed to load CLIP override '{temp_clip_name}': {e}")

        # Load VAE (Column D)
        vae_name = "Default"
        if len(found_row) > 3 and found_row[3].value:
            temp_vae_name = str(found_row[3].value).strip()
            allowed_vae = get_filename_list("vae")
            if temp_vae_name in allowed_vae:
                try:
                    vae_path = get_full_path_or_raise("vae", temp_vae_name)
                    vae = comfy.sd.load_vae(vae_path)
                    vae_name = temp_vae_name
                except Exception as e:
                    print(f"Warning: Failed to load VAE override '{temp_vae_name}': {e}")

        debug_output = f"Loadout: {loadout_name}, Model: {ckpt_name}, CLIP: {clip_name}, VAE: {vae_name}"
        
        return (model, clip, vae, debug_output)
