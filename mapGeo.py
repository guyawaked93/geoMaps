import folium
from folium import plugins
import openpyxl
from concurrent.futures import ThreadPoolExecutor

def create_marker(coord):
    popup_info = f"""
        <b>{coord['Nome da Escola']}</b><br>
        <i>{coord['Município']}</i>, {coord['UF']}<br>
        <b>Endereço:</b> {coord['Endereço']}<br>
        <b>Kit Gerador Solar:</b> {coord['Kit Gerador Solar (estimado)']}<br>
        <b>Kit Instalação Elétrica Interna:</b> {coord['Kit Instalação Elétrica Interna (estimado)']}<br>
    """
    return folium.Marker([coord["Latitude"], coord["Longitude"]], popup=folium.Popup(popup_info, max_width=300), icon=folium.Icon(color='green'))

wb = openpyxl.load_workbook("escolas.xlsx")
sheet = wb.active 

coordenadas = []
for row in sheet.iter_rows(min_row=2): 
    latitude = row[5].value
    longitude = row[6].value
    if latitude is not None and longitude is not None and isinstance(latitude, (int, float)) and isinstance(longitude, (int, float)):
        coordenadas.append({
            "UF": row[0].value,
            "Município": row[1].value,
            "Código INEP": row[2].value,
            "Nome da Escola": row[3].value,
            "Endereço": row[4].value,
            "Latitude": float(latitude),
            "Longitude": float(longitude),
            "Kit Gerador Solar (estimado)": row[7].value,
            "Kit Instalação Elétrica Interna (estimado)": row[8].value
        })

center_coords = next((coord["Latitude"], coord["Longitude"]) for coord in coordenadas 
                     if coord["Latitude"] is not None and coord["Longitude"] is not None)

m = folium.Map(location=center_coords, zoom_start=4)

cluster = plugins.MarkerCluster().add_to(m)  

with ThreadPoolExecutor() as executor:
    markers = list(executor.map(create_marker, coordenadas))

for marker in markers:
    marker.add_to(cluster)

plugins.LocateControl().add_to(m)

m.save("index.html")