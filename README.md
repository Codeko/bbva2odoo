Convierte cierto extracto bancario en (falso) formato xls en cierto extracto bancario en formato csv

## Uso:

`conversor.py input [output]`

O si hay varias versiones de python en la máquina entonces:

`python3 conversor.py input [output]`

input debe ser un archivo xls

output es opcional, y es el nombre del archivo csv de salida

Si no se define, se usa el mismo del xls pero con extensión csv

## Instalación:

Se pueden instalar las dependencias manualmente:

```shell
sudo apt install python3-pandas
sudo apt install python3-xlrd
```

O con pip3:

`pip3 install -r requeriments.txt`

En las últimas pruebas es más rápida la instalación con apt porque con pip hay que compilar pandas y tarda hasta 20 minutos.
