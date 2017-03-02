# xls
xls pipe

echo "select * from mysql.user" | mysql | ./xls.pl users.xls
