<?php

class CAQuote_Quote_Model_Observer {

    public function exportOrder(Varien_Event_Observer $observer)
	{
		$orderIds = $observer->getEvent()->getOrderIds();
		$order = Mage::getModel('sales/order')->load($orderIds[0]);
		$orderID = $order->getId();
		$this->doXML($orderID);
		
		$url = Mage::getUrl('success');
		$response = Mage::app()->getFrontController()->getResponse();
		
		$response->setRedirect($url);
		$response->sendResponse();
			
		return true;
	}
	
	public function doXML($file)
	{
		$order = Mage::getModel('sales/order')->load($file);
		#get all items
		$items = $order->getAllItems();
		$itemcount= count($items);
		$shippingDetails = $this->getOrderShippingInfo($order);
		$billingDetails = $this->getOrderBillingInfo($order);
		 
		$orderNumber = $order->getRealOrderId();
		$orderDate = $order->getCreatedAtDate();
		$orderCustomerEmail = $order->getCustomerEmail();
		$lines = $this->getOrderLineDetails($order);
		
		
		require_once dirname(__FILE__) . '/Classes/PHPExcel.php';
		
		// Create new PHPExcel object
		$objPHPExcel = new PHPExcel();
		
		// Set document properties
		$objPHPExcel->getProperties()->setCreator("WEBSITE")
									 ->setLastModifiedBy("WEBSITE.com")
									 ->setTitle("Enquiry from WEBSITE.com")
									 ->setSubject("Enquiry from WEBSITE.com")
									 ->setDescription("Enquiry from WEBSITE.com")
									 ->setKeywords("office WEBSITE php")
									 ->setCategory("WEBSITE enquiry");
		
		
		// Add some data
		$objPHPExcel->setActiveSheetIndex(0)
					->setCellValue('A1', 'Customer Information')
					->setCellValue('A3', 'Name')
					->setCellValue('B3', $billingDetails['billing_name'])
					->setCellValue('A4', 'Email')
					->setCellValue('B4', $orderCustomerEmail)
					->setCellValue('A5', 'Company Name')
					->setCellValue('B5', $billingDetails['billing_company'])
					->setCellValue('A6', 'Phone Num')
					->setCellValueExplicit('B6',$billingDetails['billing_telephone'],PHPExcel_Cell_DataType::TYPE_STRING)
					->setCellValue('A8', 'City')
					->setCellValue('B8', $billingDetails['billing_city'])
					->setCellValue('A9', 'County')
					->setCellValue('B9', $billingDetails['billing_state'])
					->setCellValue('A10', 'Postcode')
					->setCellValue('B10', $billingDetails['billing_zip'])
					->setCellValue('A13', 'Product Code')
					->setCellValue('B13', 'Description (Title)')
					->setCellValue('C13', 'Pack Size (Case Size)');

		$cellNum = "14";
		#loop for all order items
		foreach ($items as $itemId => $item)
		{
		  $cell = "A" . $cellNum;
		  $objPHPExcel->setActiveSheetIndex(0)
		  				->setCellValue($cell, $item->getSku());
		  $cell = "B" . $cellNum;
		  $objPHPExcel->setActiveSheetIndex(0)
		  				->setCellValue($cell, $item->getName());
		  $cellNum++;
		}
		
		foreach(range('A','D') as $columnID) {
    		$objPHPExcel->getActiveSheet()->getColumnDimension($columnID)
        	->setAutoSize(true);
		}
		
		$filename = "xls/WEBSITE.com_enquiry_".$file.".xlsx";
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		$objWriter->save($filename);
		
		//////////////////////////////////////////////////////////////

		include 'Mailer.php';
		
		// your email
		$recipient = "user@email.com";
		
		// person sending it
		
		$from = "noreply@email.com";
		
		// subject
		
		$subject = "New enguiry from WEBSITE.com";
		
		// email message
		
		$message = "Please see the attached quote enquiry.";

		// initialize email object ($to_address, $from_address, $subject, $reply_address=null, $mailer=null, $custom_header=null)
		
		$myEmail = new CAQuote_Mailer($recipient, $from, $subject);

		// Add the message to the email
		
		$myEmail->addText($message);



		// add the file to the email ($filename, $type=null, $filecontents=null)
		
		// NOTE: If filecontents is left out, filename is assumed to be path to file.
		
		//         If filecontents is included, filename is only used as the name of file.
		
		//            and filecontents is used as the content of file.
		
		$myEmail->addFile($filename, "application/vnd.ms-excel", $contents);
		
		
		
		// actually send out the email
		
		$myEmail->send();
		
		//////////////////////////////////////////////////////////////
		
		Mage::log('CAQuote module activated.',null,'quote.log');
	}

	public function getOrderShippingInfo($order)
	{
		$shippingAddress = !$order->getIsVirtual() ? $order->getShippingAddress() : null;
		$address_line1 = "";
		$district = "";
		
		if(strpos($shippingAddress->getData("street"), "\n")){
			$tmp = explode("\n", $shippingAddress->getData("street"));
			$district = $tmp[1];
			$address_line1 = $tmp[0];
		}
		if($address_line1 == ""){
			$address_line1 = $shippingAddress->getData("street");
		}
	 
		return array(
			 "shipping_name" =>  $shippingAddress ? $shippingAddress->getName() : '',
			 "shipping_company" =>   $shippingAddress ? $shippingAddress->getData("company") : '',
			 "shipping_street" =>    $shippingAddress ? $address_line1 : '',
			 "shipping_district" =>  $shippingAddress ? $district : '',
			 "shipping_zip" =>       $shippingAddress ? $shippingAddress->getData("postcode") : '',
			 "shipping_city" =>  $shippingAddress ? $shippingAddress->getData("city") : '',
			 "shipping_state" =>     $shippingAddress ? $shippingAddress->getRegionCode() : '',
			 "shipping_country" =>   $shippingAddress ? $shippingAddress->getCountry() : '',
			"shipping_telephone" => $shippingAddress ? $shippingAddress->getData("telephone") : ''
		);
	}
 
 
 
	public function getOrderBillingInfo($order)
	{
		$billingAddress = !$order->getIsVirtual() ? $order->getBillingAddress() : null;
		$address_line1 = "";
		$district = "";
		
		if(strpos($billingAddress->getData("street"), "\n")){
			$tmp = explode("\n", $billingAddress->getData("street"));
			$district = $tmp[1];
			$address_line1 = $tmp[0];
		}
		if($address_line1 == ""){
			$address_line1 = $billingAddress->getData("street");
		}
		return array(
			 "billing_name" =>       $billingAddress ? $billingAddress->getName() : '',
			 "billing_company" =>    $billingAddress ? $billingAddress->getData("company") : '',
			 "billing_street" =>     $billingAddress ? $address_line1 : '',
			 "billing_district" =>   $billingAddress ? $district : '',
			 "billing_zip" =>        $billingAddress ? $billingAddress->getData("postcode") : '',
			 "billing_city" =>       $billingAddress ? $billingAddress->getData("city") : '',
			 "billing_state" =>  $billingAddress ? $billingAddress->getRegionCode() : '',
			 "billing_country" =>    $billingAddress ? $billingAddress->getCountry() : '',
			"billing_telephone" =>   $billingAddress ? $billingAddress->getData("telephone") : ''
		);
	}
	 
	 
	 
	 
	public function getOrderLineDetails($order)
	{
		$lines = array();
		foreach($order->getItemsCollection() as $prod)
		{
			$line = array();
			$_product = Mage::getModel('catalog/product')->load($prod->getProductId());
			$line['sku'] = $_product->getSku();
			$line['order_quantity'] = (int)$prod->getQtyOrdered();
			$lines[] = $line;
		}
		return $lines;
	}
}
?>