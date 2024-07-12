# Use an official Python runtime as a parent image
FROM python:3.9-slim AS builder

# Set the working directory in the container
WORKDIR /app

# Install build dependencies
RUN apt-get update && apt-get install -y \
    build-essential \
    libpq-dev \
    gcc \
    && apt-get clean

# Copy the requirements file into the container
COPY requirements.txt .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Use a lightweight image for the final stage
FROM python:3.9-slim AS runtime

# Set the working directory in the container
WORKDIR /app

# Copy installed dependencies from the builder stage
COPY --from=builder /usr/local/lib/python3.9/site-packages /usr/local/lib/python3.9/site-packages
COPY --from=builder /usr/local/bin /usr/local/bin

# Copy the rest of the application code into the container
COPY . .

# Make port 80 available to the world outside this container
EXPOSE 80

# Run the application
CMD ["python", "app.py"]
