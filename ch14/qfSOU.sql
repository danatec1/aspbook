CREATE TABLE SOrder (
    SOID        int          NOT NULL,
    SOGCODE     VarChar(10)  NOT NULL,
    SOGAmount   int          NOT NULL
)

CREATE TABLE SUser (
    SUName      VarChar(10)  NOT NULL,
    SUAddress   VarChar(50)  NOT NULL,
    SOID        int          NOT NULL
)
