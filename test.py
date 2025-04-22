from PIL import Image
import os

def convertir_icône_en_noir(image_path, couleur_fond=(255, 255, 255), seuil_fond=200):
    # Ouvre l'image
    img = Image.open(image_path).convert('RGBA')  # Garder transparence
    pixels = img.load()
    largeur, hauteur = img.size

    for x in range(largeur):
        for y in range(hauteur):
            r, g, b, a = pixels[x, y]

            # Si pixel ≈ fond clair => transparent
            if r > seuil_fond and g > seuil_fond and b > seuil_fond:
                pixels[x, y] = (255, 255, 255, 0)
            else:
                # Sinon, rendre noir opaque
                pixels[x, y] = (0, 0, 0, 255)

    # Sauvegarde l’image
    img.save("edit_icon.png")

# Exemple d'utilisation
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
icon_path = os.path.join(BASE_DIR, "Icones", "edit.png")
convertir_icône_en_noir(icon_path)

