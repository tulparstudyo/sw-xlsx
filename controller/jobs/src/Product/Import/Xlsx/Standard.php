<?php

/**
 * @license LGPLv3, http://opensource.org/licenses/LGPL-3.0
 * @copyright Aimeos (aimeos.org), 2015-2020
 * @package Controller
 * @subpackage Jobs
 */

namespace Aimeos\Controller\Jobs\Product\Import\Xlsx;

/**
 * Job controller for XLSX product imports.
 *
 * @package Controller
 * @subpackage Jobs
 */
class Standard
	extends \Aimeos\Controller\Common\Product\Import\Xlsx\Base
	implements \Aimeos\Controller\Jobs\Iface
{
	private $types;


	/**
	 * Returns the localized name of the job.
	 *
	 * @return string Name of the job
	 */
	public function getName() : string
	{
		return $this->getContext()->getI18n()->dt( 'controller/jobs', 'Product import XLSX' );
	}


	/**
	 * Returns the localized description of the job.
	 *
	 * @return string Description of the job
	 */
	public function getDescription() : string
	{
		return $this->getContext()->getI18n()->dt( 'controller/jobs', 'Imports new and updates existing products from Xlsx files' );
	}


	/**
	 * Executes the job.
	 *
	 * @throws \Aimeos\Controller\Jobs\Exception If an error occurs
	 */
	public function run()
	{
		$total = $errors = 0;
		$context = $this->getContext();
		$config = $context->getConfig();
		$logger = $context->getLogger();

		if( file_exists( $config->get( 'controller/jobs/product/import/xlsx/location' ) ) === false ) {
            echo "controller/jobs/product/import/xlsx/location configuration not found in shop.php";
			return;
		}

		$domains = $config->get( 'controller/common/product/import/xlsx/domains', [] );

		$domains = $config->get( 'controller/jobs/product/import/xlsx/domains', $domains );

		$mappings = $config->get( 'controller/common/product/import/xlsx/mapping', $this->getDefaultMapping() );

		$mappings = $config->get( 'controller/jobs/product/import/xlsx/mapping', $mappings );

		$converters = $config->get( 'controller/common/product/import/xlsx/converter', [] );

		$converters = $config->get( 'controller/jobs/product/import/xlsx/converter', $converters );

		$maxcnt = (int) $config->get( 'controller/common/product/import/xlsx/max-size', 1000 );

		$skiplines = (int) $config->get( 'controller/jobs/product/import/xlsx/skip-lines', 0 );

		$strict = (bool) $config->get( 'controller/jobs/product/import/xlsx/strict', true );
        
        $backup = $config->get( 'controller/jobs/product/import/xlsx/backup' );

		if( !isset( $mappings['item'] ) || !is_array( $mappings['item'] ) )
		{
			$msg = sprintf( 'Required mapping key "%1$s" is missing or contains no array', 'item' );
			throw new \Aimeos\Controller\Jobs\Exception( $msg );
		}

		try
		{
			$procMappings = $mappings;
			unset( $procMappings['item'] );

			$codePos = $this->getCodePosition( $mappings['item'] );
			$convlist = $this->getConverterList( $converters );
			$processor = $this->getProcessors( $procMappings );
			$container = $this->getContainer();
			$msg = sprintf( 'Started product import from "%1$s" (%2$s)', $container['filename'], __CLASS__ );
			$logger->log( $msg, \Aimeos\MW\Logger\Base::NOTICE );

            $file_path = $container['location'].'/'.$container['filename'];
            $codePos = 0;

            echo "The Excel File location: $file_path\r\n";

            if(is_file($file_path)){
                $xlsx_reader = new \Aspera\Spreadsheet\XLSX\Reader();
                $xlsx_reader->open($file_path);
                $sheets = $xlsx_reader->getSheets();
                /** @var Worksheet $sheet_data */
                $line_count = 0;
                $data = [];
                foreach ($sheets as $index => $sheet_data) {
                    $xlsx_reader->changeSheet($index);
                    foreach ($xlsx_reader as $row) {
                        if($line_count++ <= $skiplines) continue;
                        if(!empty($row[$codePos])){
                            $data[$row[$codePos]] = $row;
                        }
                    }
                    break;
                }
                $xlsx_reader->close();
                echo " - ".$container['filename']." reading\r\n";
            } else{
                echo " - ".$container['filename']." not found\r\n";
                return;
            }

            $listcnt = count( $data );
            $data = $this->convertData( $convlist, $data );

            $products = $this->getProducts( array_keys( $data ), $domains );
            //echo " - Exsisting Products is getting \r\n";
            $errcnt = $this->import( $products, $data, $mappings['item'], [], $processor, $strict );
            //die('dur import');
            $chunkcnt = count( $data );

            $msg = ' - Imported product lines from "%1$s": %2$d/%3$d (%4$s)';
            $logger->log( sprintf( $msg, $container['filename'], $chunkcnt - $errcnt, $chunkcnt, __CLASS__ ), \Aimeos\MW\Logger\Base::NOTICE );

            $errors += $errcnt;
            $total += $chunkcnt;
            unset( $products, $data );
            
		}
		catch( \Exception $e )
		{
			$logger->log( 'Product import error: ' . $e->getMessage() . "\n" . $e->getTraceAsString() );
			$this->mail( 'Product XLSX import error', $e->getMessage() . "\n" . $e->getTraceAsString() );
			throw new \Aimeos\Controller\Jobs\Exception( $e->getMessage() );
		}

		$processor->finish();

		$msg = 'Finished product import from "%1$s": %2$d successful, %3$s errors, %4$s total (%5$s)';
		$logger->log( sprintf( $msg, $container['filename'], $total - $errors, $errors, $total, __CLASS__ ), \Aimeos\MW\Logger\Base::NOTICE );

		if( $errors > 0 )
		{
			$msg = sprintf( 'Invalid product lines in "%1$s": %2$d/%3$d', $container['filename'], $errors, $total );
			$this->mail( 'Product XLSX import', $msg );
			throw new \Aimeos\Controller\Jobs\Exception( $msg );
		} else {
            echo " - No errors found in $listcnt lines\r\n";
        }

		if( !empty( $backup ) && @rename( $file_path, strftime( $backup.'/'.date('Y-m-d-H-i-s').'|'.$container['filename'] ) ) === false )
		{
			$msg = sprintf( 'Unable to move imported file "%1$s" to "%2$s"', $file_path, strftime( $backup ) );
			throw new \Aimeos\Controller\Jobs\Exception( $msg );
		} else{
            echo " - ".$container['filename']." move to imported files folder \r\n";
        }
        echo " - Product importing complated\r\n";
	}


	/**
	 * Checks the given product type for validity
	 *
	 * @param string|null $type Product type or null for no type
	 * @return string New product type
	 */
	protected function checkType( string $type = null ) : string
	{
		if( !isset( $this->types ) )
		{
			$this->types = [];

			$manager = \Aimeos\MShop::create( $this->getContext(), 'product/type' );
			$search = $manager->createSearch()->setSlice( 0, 10000 );

			foreach( $manager->searchItems( $search ) as $item ) {
				$this->types[$item->getCode()] = $item->getCode();
			}
		}

		return ( isset( $this->types[$type] ) ? $this->types[$type] : 'default' );
	}


	/**
	 * Returns the position of the "product.code" column from the product item mapping
	 *
	 * @param array $mapping Mapping of the "item" columns with position as key and code as value
	 * @return int Position of the "product.code" column
	 * @throws \Aimeos\Controller\Jobs\Exception If no mapping for "product.code" is found
	 */
	protected function getCodePosition( array $mapping ) : int
	{
		foreach( $mapping as $pos => $key )
		{
			if( $key === 'product.code' ) {
				return $pos;
			}
		}

		throw new \Aimeos\Controller\Jobs\Exception( sprintf( 'No "product.code" column in XLSX mapping found' ) );
	}


	/**
	 * Opens and returns the container which includes the product data
	 *
	 * @return \Aimeos\MW\Container\Iface Container object
	 */
	protected function getContainer() : array
	{
		$config = $this->getContext()->getConfig();

		$d['location'] = $config->get( 'controller/jobs/product/import/xlsx/location' );
		
        $d['filename'] = $config->get( 'controller/jobs/product/import/xlsx/filename', 'product-import-1.xlsx' );

		$d['container'] = $config->get( 'controller/jobs/product/import/xlsx/container/type', 'Directory' );

		$d['content'] = $config->get( 'controller/jobs/product/import/xlsx/container/content', 'Binary' );

		$d['options'] = $config->get( 'controller/jobs/product/import/xlsx/container/options', [] );

		if( $d['location'] === null )
		{
			$msg = sprintf( 'Required configuration for "%1$s" is missing', 'controller/jobs/product/import/xlsx/location' );
			throw new \Aimeos\Controller\Jobs\Exception( $msg );
		}
        return $d;
		//return \Aimeos\MW\Container\Factory::getContainer( $location, $container, $content, $options );
	}


	/**
	 * Imports the XLSX data and creates new products or updates existing ones
	 *
	 * @param array $products List of products items implementing \Aimeos\MShop\Product\Item\Iface
	 * @param array $data Associative list of import data as index/value pairs
	 * @param array $mapping Associative list of positions and domain item keys
	 * @param array $types List of allowed product type codes
	 * @param \Aimeos\Controller\Common\Product\Import\Xlsx\Processor\Iface $processor Processor object
	 * @param bool $strict Log columns not mapped or silently ignore them
	 * @return int Number of products that couldn't be imported
	 * @throws \Aimeos\Controller\Jobs\Exception
	 */
	protected function import( array $products, array $data, array $mapping, array $types,
		\Aimeos\Controller\Common\Product\Import\Xlsx\Processor\Iface $processor, bool $strict ) : int
	{
		$items = [];
		$errors = 0;
		$context = $this->getContext();
		$manager = \Aimeos\MShop::create( $context, 'product' );
		$indexManager = \Aimeos\MShop::create( $context, 'index' );

		foreach( $data as $code => $list )
		{
			$manager->begin();

			try
			{
				$code = trim( $code );

				if( isset( $products[$code] ) ) {
					$product = $products[$code];
				} else {
					$product = $manager->createItem();
				}
				$map = $this->getMappedChunk( $list, $mapping );
				if( isset( $map[0] ) ) // there can only be one chunk for the base product data
				{
					$type = $this->checkType( $this->getValue( $map[0], 'product.type', $product->getType() ) );
					$map[0]['product.config'] = json_decode( $map[0]['product.config'] ?? '[]', true ) ?: [];
                    //print_r($product->toArray());
					$product = $product->fromArray( $map[0], true );
                    //print_r($product->toArray());					$product = $manager->saveItem( $product->setType( $type ) );
                    // $processor Aimeos\Controller\Common\Product\Import\Xlsx\Processor\Catalog\Standard

                    echo " -- Processor :".get_class($processor)."\r\n";
                    $list = $processor->process( $product, $list );


                    $product = $manager->saveItem( $product );
					$items[$product->getId()] = $product;
				}

				$manager->commit();
			}
			catch( \Exception $e )
			{
				$manager->rollback();

				$msg = sprintf( 'Unable to import product with code "%1$s": %2$s', $code, $e->getMessage() );
				$context->getLogger()->log( $msg );

				$errors++;
			}

			if( $strict && !empty( $list ) ) {
				$context->getLogger()->log( 'Not imported: ' . print_r( $list, true ) );
			}
		}

		$indexManager->rebuild( $items );

		return $errors;
	}
}
