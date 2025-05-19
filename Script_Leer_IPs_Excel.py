import openpyxl
import re
import ipaddress

# Archivo Excel de entrada y archivo txt de salida
archivo_entrada = "archivo_ips.xlsx"
archivo_salida = "ips_extraidas.txt"

# Regex para detectar IPs individuales y segmentos CIDR con / o |
regex_ip_cidr = re.compile(r'\b(?:\d{1,3}\.){3}\d{1,3}(?:[\/|]\d{1,2})?\b')

def normalizar_cidr(ip_cidr):
    # Cambia | por / para normalizar el formato CIDR
    return ip_cidr.replace('|', '/')

def expandir_segmento(segmento):
    try:
        # Expande el segmento a todas las IPs hosts (excluye network y broadcast)
        red = ipaddress.ip_network(segmento, strict=False)
        return [str(ip) for ip in red.hosts()]
    except ValueError:
        # No es un segmento válido, retorna como IP individual
        return [segmento]

# Cargar archivo Excel
wb = openpyxl.load_workbook(archivo_entrada, data_only=True)
ips_encontradas = []

# Recorrer todas las hojas y celdas
for hoja in wb.worksheets:
    for fila in hoja.iter_rows():
        for celda in fila:
            if celda.value:
                texto = str(celda.value)
                encontrados = regex_ip_cidr.findall(texto)
                for ip_cidr in encontrados:
                    ip_cidr_norm = normalizar_cidr(ip_cidr)
                    ips_expandidas = expandir_segmento(ip_cidr_norm)
                    ips_encontradas.extend(ips_expandidas)

# Eliminar duplicados manteniendo orden
ips_unicas = []
vistos = set()
for ip in ips_encontradas:
    if ip not in vistos:
        vistos.add(ip)
        ips_unicas.append(ip)

# Guardar todas las IPs, una por línea, en archivo txt
with open(archivo_salida, "w") as f:
    for ip in ips_unicas:
        f.write(ip + "\n")

print(f"Se extrajeron {len(ips_unicas)} IPs únicas y se guardaron en '{archivo_salida}'.")


