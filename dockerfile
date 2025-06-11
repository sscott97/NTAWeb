# Use an R-based image that includes R already
FROM rocker/r-ver:4.3.2

# Install Python and system dependencies
RUN apt-get update && \
    apt-get install -y python3 python3-pip python3-venv python3-dev && \
    apt-get install -y libcurl4-openssl-dev libssl-dev libxml2-dev

# Install any R packages your R script needs
RUN R -e "install.packages(c('ggplot2', 'dplyr'), repos='http://cran.rstudio.com/')"

# Copy and install Python dependencies
COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

# Set the working directory inside the container
WORKDIR /app

# Copy all your code files into the image
COPY . .

# Flask runs on port 10000 in this container
EXPOSE 10000

# Run your app when the container starts
CMD ["python3", "app.py"]
