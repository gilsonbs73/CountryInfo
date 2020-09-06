CREATE TABLE tb_continent(
	id int PRIMARY KEY AUTO_INCREMENT
    , continent_code char(2)
);

CREATE TABLE tb_currency(
	id int PRIMARY KEY AUTO_INCREMENT
    , currency_iso_code char(3)
);

CREATE TABLE tb_language(
	id int PRIMARY KEY AUTO_INCREMENT
    ,  language_iso_code char(2)
    ,	name_language varchar(50)
);

CREATE TABLE tb_country(
	id int PRIMARY KEY AUTO_INCREMENT
    , id_continent int REFERENCES tb_continent(id)
    , id_currency int REFERENCES tb_currency(id)
    , country_name varchar(50)
    , capital_city varchar(50)
	, phone_code char(2)
    , country_flag varchar(150)
);

CREATE TABLE tb_country_info(
	id int PRIMARY KEY AUTO_INCREMENT
    , id_country int REFERENCES tb_coutry(id)
    , id_language int REFERENCES tb_language(id)
);




