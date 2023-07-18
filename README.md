# Xls Reader
Read the Xlsx file in laravel easy with the array object 

composer require rajwebsoft/xlsreader

Add the code in config/app.php

```
'providers' => [
....
Rajwebsoft\Xlsreader\XlsreaderServiceProvider::class,
]
```

## Use in controller

```
use Rajwebsoft\Xlsreader\Xlsreader;

  $objdata =new Xlsreader();
  $xlsobj = $objdata->readFile("itf_example.xlxs");
  $sheetNames = $xlsobj->getSheetNames();
  $sheetdata=$xlsobj->getSheetData($sheetNames[1]);
  $requestcontact = $xlsobj->getDataObject($sheetdata);
```

  ## Output
```
  [{
    "name"=>"raj Kumar",
    "phone"=>"9090909090",
  },
  {
    "name"=>"rahll Kumar",
    "phone"=>"8787878788",
  }]
```
  Now you can enjoin the package
