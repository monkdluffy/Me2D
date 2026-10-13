import os
import io
import logging
from typing import Tuple
from PIL import Image

logger = logging.getLogger("ImageEngine")

class ImageEngine:
    @staticmethod
    def remove_background(input_path: str, output_path: str) -> Tuple[bool, str]:
        """
        Removes the background from an image using `rembg` and saves it as a PNG 
        with a 300 DPI resolution.
        """
        try:
            from rembg import remove, new_session
            
            if not os.path.exists(input_path):
                return False, f"Input image not found: {input_path}"
                
            img = Image.open(input_path)
            
            # To prevent memory overload and freezing, limit max resolution
            # For stamps and signatures, 2048 is more than enough
            max_size = 2048
            if max(img.size) > max_size:
                img.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
                
            # Use 'u2netp' which is the lightweight model for better performance/memory
            session = new_session("u2netp")
            
            # Perform background removal
            output_img = remove(img, session=session)
            
            # Ensure output path is .png
            if not output_path.lower().endswith('.png'):
                output_path += '.png'
                
            output_img.save(output_path, format="PNG", dpi=(300, 300))
            
            return True, f"Image saved successfully with transparent background at 300 DPI."
        except Exception as e:
            logger.error(f"Error removing background: {e}", exc_info=True)
            return False, str(e)

    @staticmethod
    def process_background_removebg_api(input_path: str, api_key: str) -> Tuple[bool, str]:
        """
        Removes the background using the remove.bg API for faster cloud processing.
        Returns a base64 encoded PNG string.
        """
        import base64
        import requests
        try:
            if not os.path.exists(input_path):
                return False, f"Input image not found: {input_path}"
                
            response = requests.post(
                'https://api.remove.bg/v1.0/removebg',
                files={'image_file': open(input_path, 'rb')},
                data={'size': 'auto'},
                headers={'X-Api-Key': api_key},
            )
            
            if response.status_code == requests.codes.ok:
                img_str = base64.b64encode(response.content).decode("utf-8")
                return True, img_str
            else:
                error_msg = f"API Error: {response.status_code} - {response.text}"
                logger.error(error_msg)
                return False, error_msg
        except Exception as e:
            logger.error(f"Error processing background via API: {e}", exc_info=True)
            return False, str(e)

    @staticmethod
    def process_background_base64(input_path: str) -> Tuple[bool, str]:
        """
        Removes the background from an image and returns a base64 encoded PNG string.
        """
        import base64
        try:
            from rembg import remove, new_session
            
            if not os.path.exists(input_path):
                return False, f"Input image not found: {input_path}"
                
            img = Image.open(input_path)
            
            max_size = 2048
            if max(img.size) > max_size:
                img.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
                
            session = new_session("u2netp")
            output_img = remove(img, session=session)
            
            buffered = io.BytesIO()
            output_img.save(buffered, format="PNG", dpi=(300, 300))
            img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
            
            return True, img_str
        except Exception as e:
            logger.error(f"Error processing background to base64: {e}", exc_info=True)
            return False, str(e)
            
    @staticmethod
    def save_base64_to_image(b64_string: str, output_path: str) -> Tuple[bool, str]:
        """
        Saves a base64 encoded image string to a file.
        """
        import base64
        try:
            if not output_path.lower().endswith('.png'):
                output_path += '.png'
            with open(output_path, "wb") as fh:
                fh.write(base64.b64decode(b64_string))
            return True, f"Image saved successfully to {output_path}"
        except Exception as e:
            logger.error(f"Error saving base64 to image: {e}", exc_info=True)
            return False, str(e)
