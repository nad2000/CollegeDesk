[![CircleCI](https://circleci.com/bb/nad2000/extracting-cell-formula-blocks-from-excel-file-and-writing-to.svg?style=svg)](https://circleci.com/bb/nad2000/extracting-cell-formula-blocks-from-excel-file-and-writing-to)


# Setup Test DB

```bash

docker rm -f mydb; docker run --name mydb -e MYSQL_ROOT_PASSWORD=p455w0rd -e MYSQL_DATABASE=blocks -p 3306:3306 -d mysql:5 --character-set-server=utf8 --collation-server=utf8_bin --default-authentication-plugin=mysql_native_password
docker logs -f mydb

# wait until the DB is up
export DATABASE_URL='mysql://root:p455w0rd@/blocks?parseTime=true'
go test ./... -short -v 
```
