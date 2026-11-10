from pathlib import Path

def get_branding(client_name: str, client_website: str = None) -> dict:

    return {
        "brand_name": "Plus94 Research",
        "primary_color": "#003366",      # Your default brand color
        "logo_path": str(Path(__file__).resolve().parent.parent / "assets" / "Plus94_logo2.png"),  # Adjust path if needed
        "client_website": client_website
    }