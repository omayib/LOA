#LOA Publisher

This Google Script will generate a PDF and send via email to participant.

Create the `Config.js` with following code

    
    const ENVIRONMENT = "prod"

    function getConfig(){
        switch(ENVIRONMENT){
        case "prod":
            return CONFIG_PRODUCTION;
        break;
        case "dev":
            return CONFIG_DEVELOPMENT;
        break;

        }
    }
    const CONFIG_PRODUCTION = {
        FORM_ID:"",
        LIST_DOSEN_ITEM_ID:"",
        LIST_MAHASISWA_ITEM_ID:"",
        SHEET_ID : "",
        DEPLOYMENT_ID:"",
        SHEET_NAME_MASTER : "",
        SHEET_NAME_DOSEN : "",
        SHEET_NAME_MHS : "",
        TEMPLATE_DOC_ID : "",
        ARCHIVE_DIR_ID :"",
        COLUMN : {
            JUDUL : 2,
            TARGET_PUBLIKASI :3,
            NAMA_MAHASISWA:4,
            EMAIL_MAHASISWA:7,
            NIM_MAHASISWA:5,
            NAMA_DOSEN:6,
            EMAIL_DOSEN:8,
            FILE_LOA:10,
            STATUS:9
        }
    }
    const CONFIG_DEVELOPMENT = {
        FORM_ID:"",
        LIST_DOSEN_ITEM_ID:"",
        LIST_MAHASISWA_ITEM_ID:"",
        SHEET_ID : "",
        DEPLOYMENT_ID:"",
        SHEET_NAME_MASTER : "",
        SHEET_NAME_DOSEN : "",
        SHEET_NAME_MHS : "",
        TEMPLATE_DOC_ID : "",
        ARCHIVE_DIR_ID :"",
        COLUMN : {
            JUDUL : 2,
            TARGET_PUBLIKASI :3,
            NAMA_MAHASISWA:4,
            EMAIL_MAHASISWA:7,
            NIM_MAHASISWA:5,
            NAMA_DOSEN:6,
            EMAIL_DOSEN:8,
            FILE_LOA:10,
            STATUS:9
        }
    }
    ```

The PDF generated from template https://docs.google.com/document/d/194yTL9P4Uwxj7ck6YdvWkMkKqxs2ox8ShHbRlQpMue4/edit?usp=sharing