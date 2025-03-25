# Usa una imagen base oficial de Python
FROM python:3.9-slim

# Configura variables de entorno para Python
ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1

# Configura el directorio de trabajo
WORKDIR /usr/src/app

# Copia los archivos de dependencias e instala paquetes
COPY requirements.txt requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copia los archivos de la aplicación
COPY . .

# Expone el puerto interno de Streamlit
EXPOSE 8501

# Comando para ejecutar Streamlit con configuraciones adicionales
CMD ["streamlit", "run", "main.py", "--server.port=8501", "--server.headless=true", "--server.enableCORS=false", "--server.enableXsrfProtection=false", "--server.address=0.0.0.0"]

