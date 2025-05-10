FROM debian:bullseye

# Instala dependências
RUN apt-get update && \
    apt-get install -y libreoffice curl unzip apache2 php php-cli php-xml php-mbstring && \
    apt-get clean

# Cria diretório para uploads
RUN rm /var/www/html/index.html

# Copia o script PHP
COPY index.php /var/www/html/index.php

# Expondo a porta padrão
EXPOSE 80

CMD ["apachectl", "-D", "FOREGROUND"]
