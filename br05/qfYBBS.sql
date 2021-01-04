CREATE TABLE ToMe (
    MID          Int          NOT NULL,
    MGrp         Int          NOT NULL,
    MSeq         Int          NOT NULL,
    MLvl         Int          NOT NULL,

    MSubject     Varchar(80)  NOT NULL,
    MContent     Text         NULL,
    MURealName   VarChar(10)  NULL,
    MUPassword   VarChar(10)  NULL,
    MVisited     Int          NULL,
    MWTime       VarChar(24)  NULL
)
