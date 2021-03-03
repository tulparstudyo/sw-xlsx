# sw-xlsx
Product Import &amp; Export for Aimeos
## console command
```
php artisan aimeos:jobs "product/export/xlsx" "default"
php artisan aimeos:jobs "product/import/xlsx" "default"
```
## Using in Laravel controller
```
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
??>
```
![image](https://user-images.githubusercontent.com/37733016/109779679-3928cf80-7c17-11eb-9913-b2b878312e1e.png)
