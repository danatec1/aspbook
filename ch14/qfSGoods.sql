CREATE TABLE SGoods (
    GCODE         VarChar(10)   NOT NULL,
    GTitle        VarChar(100)  NOT NULL,
    GPrice        int           NOT NULL,
    GDescription  text          NULL
)

INSERT INTO SGoods VALUES ('A1001','LG 386 PC', 100000,
    'LG����, CPU : 386, Memory 8M ..')
INSERT INTO SGoods  VALUES ('A1002','�ﺸ 486 PC', 150000,
    '�ﺸ ��ǻ��, CPU : 486, Memory 16M ..')
INSERT INTO SGoods VALUES ('A1003','�Ｚ 586 PC', 230000,
    '�Ｚ ����, CPU : Petium-233, Memory 32M ..')
INSERT INTO SGoods VALUES ('A1004','�Ｚ Pentium II PC', 500000,
    '�Ｚ ����, CPU : Petium II-450, Memory 64M ..')

