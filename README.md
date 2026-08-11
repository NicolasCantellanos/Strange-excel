# Strange Excel

Un juego web (novela visual, estética pixel art 16-bit) que enseña Excel desde cero a través de misiones ambientadas en el pueblo ficticio de **Vallenoche**, 1986.

🎮 **[Jugar](https://nicolascantellanos.github.io/Strange-excel/)**

## De qué se trata

Cada episodio plantea un problema con una historia y lo resuelve dentro de una hoja de cálculo real, integrada en el juego. El jugador escribe fórmulas de verdad — el juego valida el resultado, no una respuesta de memoria, y da pistas progresivas si se traba.

## Contenido (16 episodios)

| Episodio | Enseña |
|---|---|
| Entrenamiento | Celdas, filas, columnas, tu primera fórmula |
| Ep.1 | Fórmulas básicas, multiplicar, `SUMA` |
| Ep.2 | `PROMEDIO`, `MAX`, `MIN` |
| Ep.3 | `CONTAR.SI`, `SUMAR.SI` |
| Ep.4 | `BUSCARV` |
| Ep.5 | Tablas dinámicas (armadas a mano, para entender qué hacen) |
| Ep.6 | Proyecto integrador |
| Ep.7 | `SI`, `Y`, `O` (funciones lógicas) |
| Ep.8 | `MAYUSC`, `ESPACIOS`, `IZQUIERDA`, `DERECHA`, `CONCATENAR` |
| Ep.9 | `BUSCARX`, `AÑO`/`MES`/`DIA` |
| Ep.10 | Simulación de Power Query (limpieza de datos en pasos) |
| Ep.11 | `CONTAR.SI` (detectar duplicados), `VALOR` (texto a número) |
| Ep.12 | `BUSCARV` + `ESERROR` (validación de datos) |
| Ep.13 | `ÍNDICE` + `COINCIDIR` (búsqueda flexible) |
| Ep.14 | `FIN.MES`, `DIA.LAB` (fechas de negocio) |
| Ep.15 | `SI.ERROR` (auditoría de fórmulas) |
| Ep.16 | `SI` anidado (formato condicional simulado) |

## Cómo correrlo localmente

Es una app 100% estática (HTML/CSS/JS), sin build ni dependencias que instalar. Alcanza con abrir `index.html` en un navegador, o servirlo con cualquier servidor estático:

```bash
python -m http.server 8000
```

## Stack

HTML, CSS y JavaScript vanilla. La hoja de cálculo interactiva usa [jspreadsheet-ce](https://github.com/jspreadsheet/ce) vía CDN.
