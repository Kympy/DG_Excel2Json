namespace DGExcel2Json
{
    public enum EDGExcel2JsonResult
    {
        SUCCESS = 0,
        EXECUTE_ARGUMENT_REQUIRED = 1,
        EXECUTE_ARGUMENT_LENGTH = 2,
        DATA_TYPE_NOT_DEFINED = 3,
        EXCEL_NOT_EXIST = 4,
        EXCEL_PATH_WRONG = 5,
        JSON_PATH_WRONG = 6,
        SCRIPT_PATH_WRONG = 7,
        SAVE_PATH_WRONG = 8,
        NO_ID_COLUMN = 9,
        COLUMN_NAME_ERROR = 10,
        TYPE_NAME_ERROR = 11,
        DATA_READ_ERROR = 12,
        FILE_WRITE_ACCESS_DENIED = 13,
        EXCEL_IS_RUNNING = 14,
    }
}
