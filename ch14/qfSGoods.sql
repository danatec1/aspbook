CREATE TABLE SGoods (
    GCODE         VarChar(10)   NOT NULL,
    GTitle        VarChar(100)  NOT NULL,
    GPrice        int           NOT NULL,
    GDescription  text          NULL
)

INSERT INTO SGoods VALUES ('A1001','LG 386 PC', 100000,
    'LGÀüÀÚ, CPU : 386, Memory 8M ..')
INSERT INTO SGoods  VALUES ('A1002','»ïº¸ 486 PC', 150000,
    '»ïº¸ ÄÄÇ»ÅÍ, CPU : 486, Memory 16M ..')
INSERT INTO SGoods VALUES ('A1003','»ï¼º 586 PC', 230000,
    '»ï¼º ÀüÀÚ, CPU : Petium-233, Memory 32M ..')
INSERT INTO SGoods VALUES ('A1004','»ï¼º Pentium II PC', 500000,
    '»ï¼º ÀüÀÚ, CPU : Petium II-450, Memory 64M ..')

