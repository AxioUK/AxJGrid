# AxJGrid
Fork of JGrid control from J. Elihú

La idea principal es contar con un Grid con "estilo flat" que pudiera visualmente ser compatible con el LabelPlus.

The main idea is to have a "flat style" Grid that could be visually compatible with the LabelPlus.

## Caracteristicas

### Header

*	Es posible mostrar u ocultar el Header.
*	El Header permite mezclar cabeceras.
*	Es posible “skinnear” el Header con una imagen BMP, JPG
* Admite Bordes redondeados o el estilo tradicional de los grid.

### Grilla

* las celdas admiten 2 valores independientes 
```
axJGrid.CellText(Row, Col)
axJGrid.SubText(Row, Col)
```
![image](https://user-images.githubusercontent.com/61160830/118193044-20a61400-b415-11eb-9883-eba88460256e.png)

* Se pueden aplicar los siguientes colores (estos afectaran igualmente al Header con una pequeña diferencia de Alpha)
* * Bordes de Celdas
* * Color de Celdas
* * Color de Fondo (Grilla)
* * Color de Celltext
* * Color de SubText
* * Color de Selección (este color se aplica a la(s) celda(s) seleccionada(s), el color del texto cambiara automaticamente para generar un contraste y no perder visibilidad del contenido).

![image](https://user-images.githubusercontent.com/61160830/118194225-0bca8000-b417-11eb-96ce-14d056dd985e.png)



![image](https://user-images.githubusercontent.com/61160830/118166752-df513c80-b3f3-11eb-8a0d-f33475ba8bc7.png)

![image](https://user-images.githubusercontent.com/61160830/118185736-dfa90200-b40a-11eb-8fcc-f2eacb4f1e99.png)
