# Use an R-based image that includes R already
FROM rocker/r-ver:4.3.2

# Install Python and system dependencies
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        python3 python3-pip python3-venv python3-dev \
        libcurl4-openssl-dev libssl-dev libxml2-dev && \
    rm -rf /var/lib/apt/lists/*

# Install any R packages your R script needs (faster with pak)
RUN R -e "install.packages('pak', repos='https://cloud.r-project.org'); \
           pak::pak(c('cowplot', 'ggplot2', 'readxl', 'dplyr', 'tidyr'))"


# Set the working directory
WORKDIR /app

# Copy and install Python dependencies
COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

# Copy all your app code
COPY . .

# Expose Flask port
EXPOSE 10000

# Run Flask app
CMD ["python3", "app.py"]
