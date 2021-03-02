<?php

/**
 * @license LGPLv3, http://opensource.org/licenses/LGPL-3.0
 * @copyright Aimeos (aimeos.org), 2015-2020
 * @package Controller
 * @subpackage Jobs
 */ 

namespace Aimeos\Controller\Jobs\Product\Export\Xlsx;

/**
 * Job controller for product xlsx.
 *
 * @package Controller
 * @subpackage Jobs
 */
class Standard
	extends \Aimeos\Controller\Jobs\Product\Export\Standard
	implements \Aimeos\Controller\Jobs\Iface
{
	/**
	 * Returns the localized name of the job.
	 *
	 * @return string Name of the job
	 */
	public function getName() : string
	{
		return $this->getContext()->getI18n()->dt( 'controller/jobs', 'Product site map' );
	}


	/**
	 * Returns the localized description of the job.
	 *
	 * @return string Description of the job
	 */
	public function getDescription() : string
	{
		return $this->getContext()->getI18n()->dt( 'controller/jobs', 'Creates a product site map for search engines' );
	}


	/**
	 * Executes the job.
	 *
	 * @throws \Aimeos\Controller\Jobs\Exception If an error occurs
	 */
	public function run()
	{
		$container = $this->createContainer();

		$files = $this->export( $container );

		$container->close();
        return $files;
	}

	/**
	 * Adds the given products to the content object for the site map file
	 *
	 * @param \Aimeos\MW\Container\Content\Iface $content File content object
	 * @param \Aimeos\Map $items List of product items implementing \Aimeos\MShop\Product\Item\Iface
	 */
    
	protected function addItems( \Aimeos\MW\Container\Content\Iface $content, \Aimeos\Map $items )
	{
		$config = $this->getContext()->getConfig();

		$tplconf = 'controller/jobs/product/export/xlsx/standard/template-items';
		$default = 'product/export/xlsx-items-body-standard';

		$context = $this->getContext();
		$view = $context->getView();

		$view->exportItems = $items;
		$location = $config->get( 'controller/jobs/product/export/xlsx/location' );
		$view->excel_path = $location.'/'.$this->getFilename( 1 );
		$view->file_name = $this->getFilename( 1 );

		$content->add( $view->render( $context->getConfig()->get( $tplconf, $default ) ) );
	}


	/**
	 * Creates a new container for the site map file
	 *
	 * @return \Aimeos\MW\Container\Iface Container object
	 */
	protected function createContainer() : \Aimeos\MW\Container\Iface
	{
		$config = $this->getContext()->getConfig();
		$content = $config->get( 'controller/jobs/product/export/standard/container/content', 'Binary' );

		$container = $config->get( 'controller/jobs/product/export/standard/container/type', 'Directory' );

		$config = $this->getContext()->getConfig();
		$location = $config->get( 'controller/jobs/product/export/xlsx/location' );

		$options = $config->get( 'controller/jobs/product/export/standard/container/options', [] );

		if( $location === null )
		{
			$msg = sprintf( 'Required configuration for "%1$s" is missing', 'controller/jobs/product/export/location' );
			throw new \Aimeos\Controller\Jobs\Exception( $msg );
		}

		return \Aimeos\MW\Container\Factory::getContainer( $location, $container, $content, $options );
		
	}


	/**
	 * Creates a new site map content object
	 *
	 * @param \Aimeos\MW\Container\Iface $container Container object
	 * @param int $filenum New file number
	 * @return \Aimeos\MW\Container\Content\Iface New content object
	 */
	protected function createContent( \Aimeos\MW\Container\Iface $container, int $filenum ) : \Aimeos\MW\Container\Content\Iface
	{
		$config = $this->getContext()->getConfig();
		$location = $config->get( 'controller/jobs/product/export/xlsx/location' );
        if(is_file($location.'/'.$this->getFilename( $filenum )) && is_writable(dirname($location.'/'.$this->getFilename( $filenum)))){
            try{
                unlink($location.'/'.$this->getFilename( $filenum ));
            } catch(Exception $ex){
                
            }
            
        }
		
		$tplconf = 'controller/jobs/product/export/xlsx/standard/template-header';
		$default = 'product/export/xlsx-items-header-standard';

		$context = $this->getContext();
		$view = $context->getView();

		$content = $container->create( $this->getFilename( $filenum ) );
		$content->add( $view->render( $context->getConfig()->get( $tplconf, $default ) ) );
		$container->add( $content );

		return $content;
	}


	/**
	 * Closes the site map content object
	 *
	 * @param \Aimeos\MW\Container\Content\Iface $content
	 */
	protected function closeContent( \Aimeos\MW\Container\Content\Iface $content )
	{
		$tplconf = 'controller/jobs/product/export/xlsx/standard/template-footer';
		$default = 'product/export/xlsx-items-footer-standard';

		$context = $this->getContext();
		$view = $context->getView();

		$content->add( $view->render( $context->getConfig()->get( $tplconf, $default ) ) );
	}


	/**
	 * Exports the products into the given container
	 *
	 * @param \Aimeos\MW\Container\Iface $container Container object
	 * @param bool $default True to filter exported products by default criteria
	 * @return array List of content (file) names
	 */
	protected function export( \Aimeos\MW\Container\Iface $container, bool $default = true ) : array
	{
		$domains = array( 'attribute', 'media', 'price', 'product', 'text' );

		$domains = $this->getConfig( 'domains', $domains );
		$maxItems = $this->getConfig( 'max-items', 10000 );
		$maxQuery = $this->getConfig( 'max-query', 1000 );

		$start = 0; $filenum = 1;
		$names = [];

		$manager = \Aimeos\MShop::create( $this->getContext(), 'product' );

		$search = $manager->createSearch( $default );
		$search->setSortations( array( $search->sort( '+', 'product.id' ) ) );
		$search->setSlice( 0, $maxQuery );

		$content = $this->createContent( $container, $filenum );
		$names[] = $content->getResource();

		do
		{
			$items = $manager->searchItems( $search->setSlice( $start, $maxQuery ), $domains );
			$remaining = $maxItems * $filenum - $start;
			$count = count( $items );

			if( $remaining < $count )
			{
				$this->addItems( $content, $items->slice( 0, $remaining ) );
				$items = $items->slice( $remaining );

				$this->closeContent( $content );
				$content = $this->createContent( $container, ++$filenum );
				$names[] = $content->getResource();
			}

			$this->addItems( $content, $items );
			$start += $count;
		}
		while( $count >= $search->getSliceSize() );

		$this->closeContent( $content );

		return $names;
	}

	public function productlist() : array
	{
		$domains = array( 'attribute', 'media', 'price', 'product', 'text' );

		$domains = $this->getConfig( 'domains', $domains );
		$maxItems = $this->getConfig( 'max-items', 10000 );
		$maxQuery = $this->getConfig( 'max-query', 1000 );

		$start = 0; $filenum = 1;

		$manager = \Aimeos\MShop::create( $this->getContext(), 'product' );

		$search = $manager->createSearch( true );
		$search->setSortations( array( $search->sort( '+', 'product.id' ) ) );
		$search->setSlice( 0, $maxQuery );

        $items = [];
		do
		{
			$rows = $manager->searchItems( $search->setSlice( $start, $maxQuery ), $domains );
			$count = count( $rows );
			$start += $count;
            $items[] = $rows;
		}
		while( $count >= $search->getSliceSize() );

        return $items;
	}


	/**
	 * Returns the configuration value for the given name
	 *
	 * @param string $name One of "domain", "max-items" or "max-query"
	 * @param mixed $default Default value if name is unknown
	 * @return mixed Configuration value
	 */
	protected function getConfig( string $name, $default = null )
	{
		$config = $this->getContext()->getConfig();

		switch( $name )
		{
			case 'domains':
				return $config->get( 'controller/jobs/product/export/xlsx/domains', $default );

			case 'max-items':
				return $config->get( 'controller/jobs/product/export/xlsx/max-items', 50000 );

			case 'max-query':
				return $config->get( 'controller/jobs/product/export/xlsx/max-query', 1000 );
		}

		return $default;
	}


	/**
	 * Returns the file name for the new content file
	 *
	 * @param int $number Current file number
	 * @return string New file name
	 */
	public  function getXlsxFilename( int $number ) : string
	{
		return $this->getFilename($number);
	}
	protected function getFilename( int $number ) : string
	{
		return sprintf( 'product-export-%d.xlsx', $number );
	}
}
