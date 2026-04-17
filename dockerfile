# Use an R-based image that includes R already
FROM rocker/r-ver:4.3.2

# Install Python and system dependencies
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    python3 python3-pip python3-venv python3-dev \
    libcurl4-openssl-dev libssl-dev libxml2-dev && \
    rm -rf /var/lib/apt/lists/*

# Install R packages (including new ones for sigmoid fitting)
RUN R -e "install.packages('pak', repos='https://cloud.r-project.org'); \
    pak::pak(c('cowplot', 'ggplot2', 'readxl', 'dplyr', 'tidyr', 'tidyverse', 'minpack.lm'))"

# Set the working directory
WORKDIR /app

# Copy and install Python dependencies
COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

# Copy all your app code
COPY . .

# Expose Flask port
EXPOSE 10000

# Run with gunicorn (2 workers, 300s timeout for long R scripts)
CMD ["gunicorn", "--workers", "2", "--timeout", "300", "--bind", "0.0.0.0:10000", "app:app"]