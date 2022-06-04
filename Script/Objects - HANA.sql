CREATE SCHEMA "AXXIS_SAPB1";

SET SCHEMA "AXXIS_SAPB1";

DROP TABLE "AXXIS_TB_BINAutoSelect";
CREATE TABLE "AXXIS_TB_BINAutoSelect"
(
"CompanyDB" varchar(132),
"UserSign" varchar(132),
"DocNum" varchar(132),
"BINCode" varchar(132),
"RowNum" varchar(10)
);


SELECT * FROM "AXXIS_TB_BINAutoSelect";
