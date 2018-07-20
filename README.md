[![CircleCI](https://circleci.com/bb/nad2000/extracting-cell-formula-blocks-from-excel-file-and-writing-to.svg?style=svg)](https://circleci.com/bb/nad2000/extracting-cell-formula-blocks-from-excel-file-and-writing-to)


# Setup Test DB

```bash

docker rm -f mydb; docker run --name mydb -e MYSQL_ROOT_PASSWORD=p455w0rd -e MYSQL_DATABASE=blocks -p 3306:3306  mysql
```

