import os
import io
import json
import cloudinary
import cloudinary.uploader
import cloudinary.api

# Default config from environment or fallback config file
CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'static', 'cloudinary_config.json')

def load_cloudinary_config():
    """
    Loads Cloudinary configuration from environment variables or local JSON config file.
    """
    cloud_name = os.environ.get('CLOUDINARY_CLOUD_NAME', '')
    api_key = os.environ.get('CLOUDINARY_API_KEY', '')
    api_secret = os.environ.get('CLOUDINARY_API_SECRET', '')

    if not (cloud_name and api_key and api_secret) and os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
                cloud_name = saved.get('cloud_name', cloud_name)
                api_key = saved.get('api_key', api_key)
                api_secret = saved.get('api_secret', api_secret)
        except Exception:
            pass

    return {
        "cloud_name": cloud_name,
        "api_key": api_key,
        "api_secret": api_secret
    }

def init_cloudinary():
    """
    Initializes the Cloudinary SDK if credentials exist.
    """
    config = load_cloudinary_config()
    if config["cloud_name"] and config["api_key"] and config["api_secret"]:
        cloudinary.config(
            cloud_name=config["cloud_name"],
            api_key=config["api_key"],
            api_secret=config["api_secret"],
            secure=True
        )
        return True
    return False

def is_cloudinary_configured():
    config = load_cloudinary_config()
    return bool(config["cloud_name"] and config["api_key"] and config["api_secret"])

def save_cloudinary_config(cloud_name, api_key, api_secret):
    """
    Saves Cloudinary credentials to local config file and reinitializes.
    """
    data = {
        "cloud_name": cloud_name.strip(),
        "api_key": api_key.strip(),
        "api_secret": api_secret.strip()
    }
    os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)

    os.environ['CLOUDINARY_CLOUD_NAME'] = data['cloud_name']
    os.environ['CLOUDINARY_API_KEY'] = data['api_key']
    os.environ['CLOUDINARY_API_SECRET'] = data['api_secret']

    return init_cloudinary()

def upload_file_to_cloudinary(file_source, folder="msg_sheets", filename=None, resource_type="raw"):
    """
    Uploads a file (file path or BytesIO stream) to Cloudinary.
    Returns secure HTTPS URL if successful, or None if unconfigured/errored.
    """
    if not init_cloudinary():
        return None

    try:
        options = {
            "folder": f"msg_oilfield/{folder}",
            "resource_type": resource_type,
            "use_filename": True,
            "unique_filename": True
        }
        if filename:
            options["public_id"] = filename.replace(" ", "_")

        if isinstance(file_source, (str, bytes)):
            res = cloudinary.uploader.upload(file_source, **options)
        elif hasattr(file_source, 'read'):
            # BytesIO or FileStorage stream
            file_source.seek(0)
            res = cloudinary.uploader.upload(file_source, **options)
            file_source.seek(0)
        else:
            return None

        return res.get("secure_url") or res.get("url")
    except Exception as e:
        print(f"Cloudinary upload error: {e}")
        return None
