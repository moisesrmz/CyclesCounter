@echo off
echo ================================
echo 🧩 Verificando dependencias...
echo ================================

REM Lista de paquetes requeridos
set PACKAGES=watchdog matplotlib numpy

REM Función para verificar e instalar si es necesario
for %%P in (%PACKAGES%) do (
    pip show %%P >nul 2>&1
    if errorlevel 1 (
        echo 🔧 Instalando %%P...
        pip install %%P
    ) else (
        echo ✅ %%P ya está instalado.
    )
)

echo ----------------------------
echo 🚀 Todo listo para el despliegue.
pause
