# Imagen base recomendada
FROM python:3.11-slim

# Crear carpeta de trabajo
WORKDIR /app

# Copiar archivo con dependencias (si no lo tienes, te dejo requirements abajo)
COPY requirements.txt .

# Instalar dependencias
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el proyecto al contenedor
COPY . .

# Exponer puerto Flask
EXPOSE 8080

# Comando para ejecutar el servidor
CMD ["python", "app.py"]
