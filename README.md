# Generador de Documentos RRHH — v5 (LibreOffice sin depender de Debian)

⚠️ **Repo de PRUEBA, aparte de tu v4 en producción.** No reemplaces tu app
actual con esta hasta probarla tú mismo en Streamlit Cloud de verdad. La v4
sigue funcionando en su propio repo, sin tocar.

## El problema que resuelve

Streamlit Cloud corre sobre Debian 11 ("bullseye"), que llegó a su fin de
soporte oficial el 31 de agosto de 2026. Su repositorio de seguridad dejó de
renovarse, y como tu `packages.txt` le pedía instalar `libreoffice-writer`
por ese canal, tu app dejó de poder desplegarse — el error pasa ANTES de que
tu código Python siquiera se ejecute.

## Qué cambia

- **Ya no hay `packages.txt`.** Por eso ni se intenta usar `apt-get` — así
  Streamlit Cloud nunca choca con el repositorio roto de Debian.
- **LibreOffice se instala solo, la primera vez que hace falta,** descargado
  directo desde GitHub (no de Debian) en formato AppImage — un solo archivo
  ejecutable con LibreOffice completo adentro, sin necesitar `apt` para nada.
- Si en algún momento SÍ hay un LibreOffice del sistema disponible (por
  ejemplo, si terminas volviendo a usar `packages.txt` en el futuro), la app
  lo usa directo y ni se molesta en descargar nada — la descarga es solo un
  plan de respaldo automático.

## El costo real (que debes saber)

El archivo pesa ~280MB. La primera vez que la app necesita generar un PDF
después de que el servidor arrancó desde cero, hay que descargarlo y
extraerlo — eso tarda unos segundos a minutos (dependiendo de la velocidad de
red de Streamlit Cloud). Las siguientes veces, mientras el mismo servidor
siga corriendo, ya queda instalado y es instantáneo.

## Probado

- Con el sistema **sin LibreOffice instalado en absoluto** (el binario físico
  removido, simulando exactamente el escenario roto de Streamlit Cloud): la
  app lo detectó, descargó el AppImage solo, y generó 47 de 47 contratos
  reales con PDF, de principio a fin, en el navegador real.
- Declaración Jurada de Fernanda: negrita, sello, 1 página — igual que
  siempre.
- Motor de cálculo financiero (Alicia $0, Elizabeth $1.006.370): sin cambios.

---
