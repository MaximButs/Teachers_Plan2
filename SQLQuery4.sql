CREATE TABLE Plan_db (
id int IDENTITY(1,1) NOT NULL,
fio_of varchar(80) NOT NULL,
time_of int NOT NULL,
discipline varchar(50) NOT NULL,
semestr int NOT NULL,
squad varchar(50) NOT NULL,
test varchar(80) NOT NULL,
PRIMARY KEY (id)
);

create table register(
id_user int identity(1,1) NOT NULL,
login_user varchar(50) NOT NULL,
password_user varchar(50) NOT NULL,
is_admin bit
);

--select * from register

insert into register (login_user, password_user, is_admin) values ('admin', 'asd', 1)