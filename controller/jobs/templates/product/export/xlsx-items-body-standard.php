<?php
$enc = $this->encoder();
$domains = array( 'attribute', 'media', 'price', 'product', 'text' );

$cols = ["item_code" => 'string', 
		 "item_label" => 'string', 
		 "item_type" => 'string', 
		 "item_status" =>'string', 
		 "text_type" => 'string', 
		 "text_content" => 'string', 
		 "text_type_2" => 'string', 
		 "text_content_2" => 'string', 
		 "media_url" => 'string', 
		 "price_currencyid" =>'string', 
		 "price_quantity" => 'string', 
		 "price_value" => 'string', 
		 "price_tax_rate" => 'string', 
		 "attribute_code" => 'string', 
		 "attribute_type" => 'string', 
		 "subproduct _code" => 'string', 
		 "product_list type" => 'string', 
		 "property_value" => 'string', 
		 "property_type" => 'string', 
		 "catalog_code" => 'string', 
		 "catalog_list_type" => 'string', 
		 "media_url2" => 'string', 
		 "media_url3" => 'string', 

		];

$wExcel = new Ellumilel\ExcelWriter();
$wExcel->writeSheetHeader('Sheet1', $cols);
$wExcel->setAuthor('Swordbros.ru');

foreach( $this->get( 'exportItems', [] ) as $item ){
    $media_url = '';
    $price_currencyid = '';
    $price_value = '0';
    $text_type = 'short';
    $text_content = '';
    $text_type_2 = 'long';
    $text_content_2 = '';
    $price_tax_rate = '0';
    $media_urls = [];
    foreach( $domains as $domain ){
            if($listItems = $item->getListItems( $domain )){
                if($domain=='media'){
                    foreach($listItems as $listItem){
                        $info = $listItem->getRefItem();
                        if(!in_array($info->get('media.url'), $media_urls))$media_urls[] = $info->get('media.url');
                    }
                } if($domain=='price'){
                    if($listItem = $listItems->first()){
                        $info = $listItem->getRefItem();
                        $price_currencyid = ($info->get('price.currencyid'));
                        $price_value = ($info->get('price.value'));
                    }
                } if($domain=='text'){
                    foreach($listItems as $listItem ){
                        $info = $listItem->getRefItem();
                        if($info->get('text.type')=='short'){
                            $text_type = 'short';
                            $text_content = $info->get('text.content');
                        } elseif($info->get('text.type')=='long'){
                            $text_type_2 = 'long';
                            $text_content_2 = $info->get('text.content');
                        }
                    }
                }
            }


    }

    $label = $item->getLabel();
    $label = preg_replace('/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]/u', '', $label);
    $validUTF8 = mb_check_encoding($label, 'UTF-8');
    if( $validUTF8 && $item->getCode() && $item->getLabel() && $item->getType( ) && $price_currencyid){
        $row = ["item_code" => $enc->xml( $item->getCode() ), 
                 "item_label" => $enc->xml( $label ),
                 "item_type" => $enc->xml( $item->getType() ),
                 "item_status" => $enc->xml( $item->getStatus() ),
                 "text_type" => $text_type,
                 "text_content" =>  preg_replace( "/\r|\n/", "", $text_content ) ,
                 "text_type_2" => $text_type_2,
                 "text_content_2" => preg_replace( "/\r|\n/", "", $text_content_2 ) ,
                 "media_url" => isset($media_urls[0])?$media_urls[0]:'',
                 "price_currencyid" => $price_currencyid ,
                 "price_quantity" => '1',
                 "price_value" => $price_value,
                 "price_tax_rate" => $price_tax_rate,
                 "attribute_code" => '',
                 "attribute_type" => '',
                 "subproduct_code" => '',
                 "product_list_type" => '',
                 "property_value" => '',
                 "property_type" => '',
                 "catalog_code" => '',
                 "catalog_list_type" => 'standard',
                 "media_url2" => isset($media_urls[1])?$media_urls[1]:'',
                 "media_url3" => isset($media_urls[2])?$media_urls[2]:'',
            ];
        $wExcel->writeSheetRow('Sheet1', $row);	
    }    

} 
   
$excel_path = $this->get('excel_path');

if(is_file($excel_path)){
    unlink($excel_path);
}
$wExcel->writeToFile($excel_path); 
//echo "Yazıldı: $excel_path";
die();