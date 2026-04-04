# Gunakan image Node.js versi LTS (Alpine untuk ukuran lebih kecil dan aman)
FROM node:20-alpine

# Update paket di alpine dan instal openssl (dibutuhkan oleh Prisma)
RUN apk add --no-cache openssl

# Set working directory di dalam container
WORKDIR /usr/src/app

# Salin package.json dan package-lock.json
COPY package*.json ./

# Salin direktori prisma (untuk generate Prisma Client)
COPY prisma ./prisma/

# Install seluruh dependency (termasuk devDependencies karena butuh CLI 'prisma' untuk generate)
# Hapus --production agar devDependencies ikut terinstall untuk 'prisma generate' saat postinstall
RUN npm install

# Salin seluruh kode aplikasi ke dalam container
COPY . .

# Expose port (Cloud Run akan menggunakan environment variable PORT, dan by default 8080)
EXPOSE 8080

# Menyiapkan environment variable default jika tidak ada (Cloud Run akan timpa ini jika diset di console)
ENV PORT=8080
ENV NODE_ENV=production

# Perintah untuk menjalankan server
CMD ["node", "server.js"]
