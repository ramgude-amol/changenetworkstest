<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Test extends CI_Controller 
{
	
	function __construct() 
    {
		
        parent::__construct();
    }
	public function index()
	{
		$data['error'] = "";
		if(!empty($_FILES['upload'])){
			$status = $this->uploads();
			$data['error'] = !$status ? "One or more files uploaded are not with doc or docx extension" : '';
			if($status === true){
				redirect("/test/lists");
			}
		}
		 $this->load->view('test/index',$data);
	}
	
	public function uploads()
	{
		$path = ASSETSPATH. DIRECTORY_SEPARATOR. "docs";
		
		$total = count($_FILES['upload']['name']);

		// Loop through each file
		for( $i=0 ; $i < $total ; $i++ ) 
		{
			 $filename = $_FILES['upload']['name'][$i];

			$ext = pathinfo($filename, PATHINFO_EXTENSION);
			if ($ext !== 'doc' && $ext !== 'docx') {
				return false;
			}
		
			//Get the temp file path
			$tmpFilePath = $_FILES['upload']['tmp_name'][$i];

			//Make sure we have a file path
			if ($tmpFilePath != ""){
				//Setup our new file path
				$filename ="FILE_".time().".doc";
				$newFilePath = $path.DIRECTORY_SEPARATOR. $filename;

				//Upload the file into the temp dir
				if(!move_uploaded_file($tmpFilePath, $newFilePath)) {
					return false;
				}
				
			    $pdf_path = ASSETSPATH. DIRECTORY_SEPARATOR. "pdfs";
				$pdf_path = $pdf_path. DIRECTORY_SEPARATOR. (str_replace([".doc",".docx"],"",$filename)).".pdf";
				$this->word2pdf($newFilePath,$pdf_path);
			}
		}
		return true;
	}
	
	
   function word2pdf($lastfnamedoc,$lastfnamepdf)
	{ 
		try{
		   $word = new COM("Word.Application") or die ("Could not initialise Object.");
		   // set it to 1 to see the MS Word window (the actual opening of the document)
		   $word->Visible = 1;
		   // recommend to set to 0, disables alerts like "Do you want MS Word to be the default .. etc"
		   $word->DisplayAlerts = 1;
		   // open the word 2007-2013 document 
		   //$word->Documents->Open('C:\xampp\htdocs\Open_Office\test_4sk.docx');
		   // save it as word 2003
		   $word->Documents->Open($lastfnamedoc);
		   // convert word 2007-2013 to PDF
		   $word->ActiveDocument->ExportAsFixedFormat($lastfnamepdf, 17, false, 0, 0, 0, 0, 7, true, true, 2, true, true, false);
		   // quit the Word process
		   $word->Quit(false);
		   // clean up
		   unset($word);
	   }catch(Exception $e){
		   echo $e;
	   }
	}	
	
	
	public function lists()
	{
		
		$path = ASSETSPATH. DIRECTORY_SEPARATOR . 'pdfs';
		$data = ['path'=>$path];
		$files = scandir($path);
		if(empty($files))
		{
			redirect('test/index');
		}
		rsort($files); 
		
		foreach($files as $file)
		{
		  $ext = pathinfo($file, PATHINFO_EXTENSION);
		  if($ext == 'pdf'){
			  $data['files'][] = $file;
		  }
		}
		$this->load->view('test/lists',$data);
	}
}
