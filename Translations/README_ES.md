# PS-NCDU

**Analizador de espacio en disco para Windows, en PowerShell, con interfaz web local y árbol navegable en tiempo real.**

[![Version](https://img.shields.io/badge/version-6.28-2c6cb0)](https://github.com/Vietnamix/PS-NCDU)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?logo=powershell&logoColor=white)](https://github.com/Vietnamix/PS-NCDU)
[![Plataforma](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white)](https://github.com/Vietnamix/PS-NCDU)
[![Licencia](https://img.shields.io/badge/license-MIT-3fa45b)](License.md)

PS-NCDU es un script de PowerShell **de un solo archivo y autónomo** que responde en segundos a la pregunta «¿qué está llenando este disco?». Ejecutado sin ningún parámetro, arranca un pequeño servidor web local, abre el navegador y le permite elegir una carpeta o un disco para analizar. El árbol se construye **en directo** durante el escaneo, los tamaños se rellenan carpeta por carpeta, y puede navegar libremente incluso antes de que termine. Inspirado en la herramienta Unix [`ncdu`](https://dev.yorhel.nl/ncdu), pensado para el ecosistema Windows, sin ninguna dependencia externa.

![Interfaz de PS-NCDU](PS-NCDU_interface_v6.27b.png)

*PowerShell 5.1+ · Windows · Sin instalación · Un solo archivo · 42 idiomas*

Otros idiomas: [Français](README.md) · [English](README_EN.md) · [中文](README_ZH.md) · [हिन्दी](README_HI.md) · [العربية](README_AR.md) · [বাংলা](README_BN.md) · [Português](README_PT.md) · [Русский](README_RU.md) · [اردو](README_UR.md) · [Bahasa Indonesia](README_ID.md) · [Deutsch](README_DE.md) · [日本語](README_JA.md) · [Türkçe](README_TR.md) · [Tiếng Việt](README_VI.md) · [한국어](README_KO.md) · [Italiano](README_IT.md)

---

## Índice

- [Descripción general](#descripción-general)
- [Funcionalidades](#funcionalidades)
- [Requisitos](#requisitos)
- [Instalación](#instalación)
- [Uso](#uso)
- [La interfaz](#la-interfaz)
- [Cómo funciona el escaneo](#cómo-funciona-el-escaneo)
- [Archivos de trabajo](#archivos-de-trabajo)
- [Solución de problemas](#solución-de-problemas)
- [Hoja de ruta](#hoja-de-ruta)
- [Contribuir](#contribuir)
- [Licencia](#licencia)

---

## Descripción general

A diferencia de las versiones 3.x, que producían un informe HTML estático para abrir después, PS-NCDU es ahora una **aplicación web local**. El script arranca un servidor HTTP en `127.0.0.1` (puerto 8787, con respaldo automático a un puerto libre), protegido por un token de sesión, y abre la interfaz en su navegador predeterminado. Todo se configura desde esa interfaz: la carpeta a analizar, la profundidad, el filtro de visualización, las exclusiones, el idioma.

El servidor permanece en la máquina local, no se expone en la red, y se detiene con `Ctrl+C` en la consola o con el botón «Cerrar el servidor» de la interfaz.

![Ventana de análisis de PS-NCDU](PS-NCDU_scan_form_v6.27b.png)

---

## Funcionalidades

### Escaneo y navegación
- **Árbol en tiempo real**: la estructura aparece durante la enumeración, y los tamaños llegan carpeta por carpeta a medida que se calculan.
- **Navegación libre durante el escaneo**: haga clic en una carpeta para entrar, use la ruta de navegación para subir, sin esperar al final.
- **Profundidad ajustable o ilimitada**: precargue unos niveles para una visualización fluida, o el árbol completo. Los tamaños son siempre exactos sea cual sea la profundidad; los niveles no precargados se cargan con un clic.
- **Escaneo ilimitado intercalado**: en profundidad ilimitada, cada subárbol se enumera justo antes de medirse, de modo que los tamaños aparecen en los primeros segundos en lugar de esperar a recorrer todo el disco.
- **Interrupción**: un botón «Detener» para el escaneo en curso y devuelve el control; iniciar un nuevo escaneo cancela automáticamente el anterior.
- **Carpetas grises clicables**: excluidas, uniones, protegidas (ACL) o no precargadas permanecen visibles y se escanean bajo demanda, con una cola si ya hay un escaneo en marcha.
- **Orden de procesamiento alineado con la visualización**: el progreso se rellena de arriba abajo, sin saltos.

### Lectura de resultados
- **Orden Nombre / Tamaño**: por tamaño de forma predeterminada (los mayores arriba), suavizado durante el escaneo para que las filas no salten; un clic cambia al orden por nombre.
- **Barras de proporción** y porcentajes respecto a la carpeta actual.
- **Contadores recursivos** de subcarpetas y archivos por carpeta.
- **Archivos listados bajo demanda** al abrir una carpeta, ordenados por tamaño, limitados a los 1000 mayores.
- **Iconos por tipo de archivo**: unas 120 extensiones comunes (imágenes, vídeo, audio, PDF, ofimática, archivos comprimidos, código, ejecutables, fuentes, imágenes de disco, bases de datos, ebooks, certificados, accesos directos) para identificar tipos de un vistazo.
- **Filtro de visualización**: ocultar elementos menores de 1 MB, 100 MB o 1 GB para mayor legibilidad, sin modificar el escaneo.
- **Puntos de estado de colores**: escaneado, previsto, en cola, en curso, gris; una leyenda y una ayuda integradas explican cada estado.
- **Tema claro / oscuro**.

### Ventana de análisis
- **Explorador de carpetas integrado**: unidades, navegación por clic, carpeta superior, «Elegir esta carpeta». Sin dependencia del selector nativo de Windows, por lo que es fiable incluso en acceso remoto.
- **Accesos rápidos** a los perfiles de usuario, **recientes** con eliminación individual y borrado, **unidades** con barra de ocupación.
- **Validación de la ruta** en directo y al iniciar.
- **Exclusiones**: lista de carpetas del sistema siempre omitidas, más un campo para excluir otras durante un escaneo.
- **Ajustes recordados** (ruta, profundidad, filtro, orden, idioma) entre sesiones.
- **Intro** para iniciar, **Esc** o el botón de cierre para ocultar la ventana cuando ya se muestra un escaneo.

### Idiomas
- **42 idiomas**, que cubren más del 80 % de la población mundial: inglés, chino, hindi, español, francés, árabe, bengalí, portugués, ruso, urdu, indonesio, alemán, japonés, coreano, italiano, turco, vietnamita, polaco, neerlandés, ucraniano, rumano, checo, griego, sueco, húngaro, persa, tailandés, malayo, filipino, suajili, tamil, telugu, maratí, guyaratí, canarés, malayalam, panyabí, hebreo, hausa, birmano, amárico, jemer.
- Detección automática del idioma del sistema, selector en la ventana de análisis, elección recordada.
- Escritura de derecha a izquierda para árabe, urdu, persa y hebreo.
- Las traducciones de los 30 idiomas más recientes son de mejor esfuerzo; se agradece la revisión por hablantes nativos, en especial para amárico, jemer, birmano, hausa y las lenguas de la India.

---

## Requisitos

| Elemento   | Detalle                                                                                                                             |
| ---------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| Sistema    | Windows 10 / 11 o Windows Server                                                                                                    |
| PowerShell | 5.1 (Windows PowerShell) o 7+ (PowerShell Core)                                                                                     |
| Modo       | **FullLanguage** requerido (el servidor web se apoya en `HttpListener`). Vea [Solución de problemas](#solución-de-problemas) para el modo restringido. |
| Permisos   | Lectura sobre las carpetas escaneadas; algunas rutas del sistema exigen una consola de **administrador**                            |
| Navegador  | Cualquier navegador reciente                                                                                                        |

No se requiere ningún módulo externo.

---

## Instalación

Clone el repositorio o simplemente descargue el archivo `ps-ncdu.ps1`:

```powershell
git clone https://github.com/Vietnamix/PS-NCDU.git
cd PS-NCDU
```

El script está codificado en **UTF-8 con BOM**. No lo vuelva a guardar en otra codificación: PowerShell 5.1 leería entonces el archivo como ANSI y rompería los acentos y los idiomas no latinos de la interfaz.

> **Política de ejecución**: si Windows bloquea la ejecución de scripts, autorícela para la sesión actual:
>
> ```powershell
> Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
> ```
>
> Este comando no cambia nada de forma permanente: solo vale para la ventana de PowerShell abierta.

---

## Uso

**No hay parámetros de línea de comandos**. Simplemente ejecute el script:

```powershell
.\ps-ncdu.ps1
```

El script:

1. arranca el servidor local y muestra su dirección en la consola (por ejemplo `http://127.0.0.1:8787/?token=...`);
2. abre esa dirección en su navegador predeterminado;
3. presenta la ventana de análisis, donde elige la carpeta, la profundidad y la visualización, y luego hace clic en «Analizar».

Si el navegador no se abre por sí solo, copie la dirección mostrada en la consola. Para detener: `Ctrl+C` en la consola, o el botón «Cerrar el servidor» de la interfaz.

Para escanear rutas del sistema (`C:\Windows`, la raíz de un disco), inicie la consola **como administrador**: de lo contrario las carpetas protegidas aparecerán en gris.

---

## La interfaz

- **Cabecera**: ruta de navegación, total de la carpeta actual con contadores, porcentaje de avance, botón de orden Nombre / Tamaño, tema, «Detener» durante un escaneo, «Nuevo escaneo».
- **Árbol**: una fila por carpeta o archivo, con punto de estado, icono de tipo, nombre, barra de proporción, porcentaje, contadores y tamaño. Las carpetas grises se escanean con un clic; un botón «Escanear carpetas grises (N)» procesa todas las de la carpeta actual.
- **Pie de página**: etapa en curso, ruta que se está leyendo en este instante, profundidad del escaneo, cronómetro, y el panel Leyenda / Ayuda.
- **Ventana de análisis** («Nuevo escaneo»): en dos columnas. A la izquierda el destino (ruta, explorador integrado, accesos rápidos, recientes, unidades). A la derecha las opciones (profundidad con deslizador y modo ilimitado, filtro de visualización, exclusiones). El selector de idioma está en la cabecera de esta ventana.

---

## Cómo funciona el escaneo

El motor trabaja por etapas. Una **enumeración** descubre la estructura y la emite al árbol sobre la marcha; un **cálculo de tamaños** recorre después cada subárbol de primer nivel, en el orden de visualización, subiendo los tamaños parciales a todos los ancestros una vez por segundo. Los eventos fluyen del servidor a la página mediante SSE.

Algunas decisiones de diseño que conviene conocer:

- **Un solo escaneo a la vez**, deliberadamente: dos escaneos de disco en paralelo se ralentizarían mutuamente. Las solicitudes adicionales (carpetas grises) se ponen en cola y se procesan al final.
- **Servidor de un solo hilo**: la interrupción es cooperativa. Cerrar la conexión (botón «Detener», o nuevo escaneo) hace fallar la siguiente escritura del servidor, lo que activa una bandera comprobada en los bucles; la parada efectiva tarda hasta un segundo.
- **Uniones y puntos de reanálisis** se omiten para evitar bucles y dobles conteos.
- **Unidades de red**: su espacio no se consulta al arrancar, lo que evita un bloqueo cuando una unidad mapeada es inaccesible (VPN caída, por ejemplo).
- **Enumeración .NET en flujo** (`EnumerateFiles` / `EnumerateDirectories`) en lugar de `Get-ChildItem`, notablemente más rápida en PowerShell 5.1. El techo de rendimiento sigue siendo el de un script interpretado: las herramientas nativas que leen la MFT de NTFS directamente son mucho más rápidas, y eso se asume.

---

## Archivos de trabajo

| Ubicación                            | Función                                        |
| ------------------------------------ | ---------------------------------------------- |
| `%TEMP%\psncdu\psncdu_debug.log`     | Registro detallado del servidor y los escaneos |
| `%TEMP%\psncdu\history.txt`          | Historial de rutas escaneadas («Recientes»)     |

Carpetas del sistema siempre excluidas del escaneo: `C:\Windows\WinSxS`, `C:\Windows\Installer`, `C:\$Recycle.Bin`, `C:\System Volume Information`, `C:\Recovery`, `C:\ProgramData\Microsoft\Windows Defender`, `C:\Windows\SoftwareDistribution`. Puede añadir otras, durante un escaneo, desde la sección Exclusiones de la ventana de análisis.

---

## Solución de problemas

| Síntoma                                                     | Causa probable / solución                                                                                                                       |
| ----------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| El script no arranca                                        | Política de ejecución, vea la nota en [Instalación](#instalación).                                                                              |
| «El servidor web exige el modo FullLanguage»                | La sesión está en ConstrainedLanguage (política AppLocker / WDAC). Ejecute desde una consola no restringida, o use la versión 3.2 (informe HTML estático), que funciona en modo restringido. |
| Acentos o idiomas no latinos rotos en la interfaz           | El `.ps1` se volvió a guardar sin BOM UTF-8. Restaure la codificación original.                                                                  |
| Muchas carpetas grises «protegido»                          | Permisos insuficientes. Reinicie PowerShell **como administrador**.                                                                             |
| El navegador no se abre                                     | Abra manualmente la dirección mostrada en la consola (con su token).                                                                            |
| Página en blanco o servidor bloqueado al arrancar           | Revise el registro `%TEMP%\psncdu\psncdu_debug.log`. Una unidad de red inaccesible podía bloquear versiones antiguas; corregido desde la 5.14.  |
| No pasa nada durante un escaneo ilimitado de un disco       | Corregido desde la 6.17 (enumeración intercalada). Si persiste, compruebe la versión mostrada en la cabecera.                                    |

---

## Hoja de ruta

- [ ] Revisión de las traducciones de los 30 idiomas recientes por hablantes nativos
- [ ] Barra de progreso anclada al espacio realmente ocupado del disco, en lugar de a tramos de etapas
- [ ] Indicadores de rendimiento durante el escaneo (archivos por segundo, MB por segundo)
- [ ] Exportación CSV / JSON de los resultados
- [ ] Comparación de dos escaneos a lo largo del tiempo

*Sugerencias bienvenidas a través de issues.*

---

## Contribuir

Las contribuciones son bienvenidas:

1. Haga *fork* del repositorio.
2. Cree una rama (`git checkout -b feature/mi-funcionalidad`).
3. Conserve la codificación **UTF-8 con BOM** y los *here-strings* de PowerShell intactos.
4. Para las traducciones, cada idioma es un objeto del diccionario `I18N` dentro del script; compare sus claves con las de `en` para detectar lo que falta.
5. Abra una *pull request* describiendo claramente el cambio.

Para errores e ideas, abra un **issue** indicando la versión de Windows, la versión de PowerShell, la versión de PS-NCDU mostrada en la cabecera y, si es posible, un extracto del registro `%TEMP%\psncdu\psncdu_debug.log`.

---

## Licencia

Distribuido bajo licencia **MIT**. Vea el archivo [`License.md`](License.md).

---

## Autor

**[Eric Guiffault](https://eric.guiffault.com)**

Si este proyecto le resulta útil, considere dejar una estrella en GitHub.
