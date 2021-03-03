# sw-xlsx
Product Import &amp; Export for Aimeos
## console command
```
php artisan aimeos:jobs "product/export/xlsx" "default"
php artisan aimeos:jobs "product/import/xlsx" "default"
```
## Example of using Laravel on the controller
```
// Export xlsx file
<?php
$Xlsx = new \Aimeos\Controller\Jobs\Product\Export\Xlsx\Standard($this->get('context'), $this->get('aimeos'));
$files = $Xlsx->run();
if($files){
   $file_path = $files[0];
    header('Content-Type: application/download');
    header("Content-Disposition: attachment; filename=".$Xlsx->getXlsxFilename(1)); 
    header("Pragma: no-cache"); 
    header("Expires: 0"); 
    readfile($file_path);
    die();
}
?>

// Import xlsx file
$location = $config->get( 'client/product/import/xlsx/location' );
$file_name = 'product-import-1.xlsx';
$file_path = $location .'/'.$file_name;

if($files = (array) $this->getView()->request()->getUploadedFiles()){
    foreach($files as $file ){
        $file->moveTo($file_path);
    }
    $Xlsx = new \Aimeos\Controller\Jobs\Product\Import\Xlsx\Standard($context, $this->getAimeos());
    $Xlsx->run();
}
$context->getSession()->set( 'info', [$context->getI18n()->dt( 'admin', 'Items imported successfully' )] ); 
return $this->redirect( 'product', 'search');

```
![image](https://user-images.githubusercontent.com/37733016/109779679-3928cf80-7c17-11eb-9913-b2b878312e1e.png)
