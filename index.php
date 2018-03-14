<?php 
    //////////////////////////////////////////////// 
    // PHP Office VBA password remover. 
    // 
    // Created by /u/AyrA_ch for /r/excel 
    // Feel free to do whatever you want 
    // as long as access to this stays free of 
    // any charges and/or subscriptions. 
    // 
    // Working copy at https://home.ayra.ch/unlock/ 
    // 
    // You need the bootstrap framework for this to look nice. 
    // It's not required for the functionality itself. 
    // 
    /////////////////////////////////////////////// 
     
    //Be extra pedantic 
    error_reporting(E_ALL); 
     
    $err=""; 
     
    //Extract vbaProject.bin from a modern office file. 
    function getFromZip($fName){ 
        global $err; 
        $temp=""; 
        $zip=new ZipArchive(); 
         
        //Try to open the file as zip archive 
        if($res=$zip->open($fName) && $zip->numFiles>0){ 
            //try to figure out, where to extract vbaProject.bin from (excel,word,powerpoint) 
            if(($temp=$zip->getFromName("xl/vbaProject.bin"))==FALSE){ 
                if(($temp=$zip->getFromName("word/vbaProject.bin"))==FALSE){ 
                    if(($temp=$zip->getFromName("ppt/vbaProject.bin"))==FALSE){ 
                        $err="No VBA Code in the specified file. This script only removes VBA protection and not document protection (password to open or change content)";
                    } 
                } 
            } 
            $zip->close(); 
        } 
        else{ 
            $err="Can't open your file. It seems to be from Office 2007 or newer but is corrupt."; 
        } 
        //return the vbaProject.bin content if it has been extracted. 
        return $temp===FALSE?"":$temp; 
    } 

    //Add vbaProject.bin back to a modern Office file 
    function addToZip($contents,$fName) 
    { 
        global $err; 
        $temp=""; 
        $zip=new ZipArchive; 
        //Open file as zip archive 
        if($res=$zip->open($fName)) 
        { 
            //Try to find where the original vbaProject.bin was located and replace it with decrypted blob. 
            if($zip->getFromName("xl/vbaProject.bin")==FALSE) 
            { 
                if($temp=$zip->getFromName("word/vbaProject.bin")==FALSE){ 
                    //Powerpoint 
                    $zip->deleteName("ppt/vbaProject.bin"); 
                    $zip->addFromString("ppt/vbaProject.bin",$contents); 
                } 
                else{ 
                    //Word 
                    $zip->deleteName("word/vbaProject.bin"); 
                    $zip->addFromString("word/vbaProject.bin",$contents); 
                } 
            } 
            else{ 
                //Excel 
                $zip->deleteName("xl/vbaProject.bin"); 
                $zip->addFromString("xl/vbaProject.bin",$contents); 
            } 
            $zip->close(); 
        } 
        else{ 
            $err="Can't open file to change VBA settings. Inform the administrator."; 
        } 
    } 
     
    //provides file source code upon request 
    if(isset($_GET["source"])){ 
        echo '<html><head><meta http-equiv="X-UA-Compatible" content="IE=edge" /><meta name="viewport" content="width=device-width, initial-scale=1" /><title>VBA Unlocker Source</title></head><body>'; 
        highlight_file(__FILE__); 
        echo "</body></html>"; 
        exit(0); 
    } 
     
    //Check if file uploaded 
    if(isset($_FILES['excel'])){ 
        //Ensure a file name exists 
        if(isset($_FILES['excel']['tmp_name']) && $_FILES['excel']['tmp_name']!=""){ 
            //Try to read the file 
            if($fp=fopen($_FILES['excel']['tmp_name'],"rb")){ 
                //Read everything and close file handle 
                $contents=fread($fp,filesize($_FILES['excel']['tmp_name'])); 
                fclose($fp); 

                //If it starts with "PK" it is a modern file (O2007 and newer) 
                if(substr($contents,0,2)=="PK"){ 
                    //Create temporary zip file name 
                    $z=tempnam(dirname(__FILE__)."/TMP/","zip"); 
                    //Move temporary file to place where it's not readonly to us. 
                    move_uploaded_file($_FILES['excel']['tmp_name'], $z); 
                    //Get VBA blob from zip 
                    $contents=getFromZip($z); 
                    if($contents!="" && $err==""){ 
                        if(strpos($contents,"DPB=")===FALSE){ 
                            $err="We found VBA Code but it is not protected."; 
                        } 
                        else{ 
                            $contents=str_replace("DPB=","DPx=",$contents); 
                            addToZip($contents,$z); 
                            if($err=="") 
                            { 
                                if($fp=fopen($z,"rb")){ 
                                    header("Content-Type: application/octet-stream"); 
                                    header("Content-Disposition: attachment; filename=\"" . $_FILES['excel']['name'] . "\""); 
                                    echo fread($fp,filesize($z)); 
                                    fclose($fp); 
                                    //Delete the uploaded file on success 
                                    unlink($z); 
                                    exit(0); 
                                } 
                                else{ 
                                    $err="Can't send back office file. Error opening the temporary file."; 
                                } 
                            } 
                        } 
                    } 
                    else{ 
                        $err="This file has no encrypted VBA code, or the entire file is encrypted."; 
                    } 
                    //Delete the uploaded file on error 
                    unlink($z); 
                } 
                else{ 
                    //Delete uploaded file because it's in $contents now 
                    unlink($_FILES['excel']['tmp_name']); 
                     
                    //assume classic file (O2003 and older) 
                    if(strpos($contents,"DPB=")===FALSE){ 
                        $err="There is no VBA code or it is not protected."; 
                    } 
                    else{ 
                        //This removes the protection 
                        $contents=str_replace("DPB=","DPx=",$contents); 
                        //Send back file 
                        header("Content-Disposition: attachment; filename=\"" . $_FILES['excel']['name'] . "\""); 
                        header("Content-Type: application/octet-stream"); 
                        echo $contents; 
                        exit(0); 
                    } 
                } 
            } 
            else{ 
                $err="We were unable to open the file. Either our disk is full or it was removed by our Anti-Virus"; 
            } 
        } 
        else{ 
            $err="No file received. Please select a file to decrypt"; 
        } 
    } 
?><!DOCTYPE html> 
<html> 
    <head> 
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" /> 
        <meta http-equiv="X-UA-Compatible" content="IE=edge" /> 
        <meta name="viewport" content="width=device-width, initial-scale=1" /> 
        <link rel="stylesheet" type="text/css" href="../bootstrap/css/bootstrap.min.css" /> 
        <link rel="stylesheet" type="text/css" href="../bootstrap/css/bootstrap-theme.min.css" /> 
        <script type="text/javascript" src="../bootstrap/js/jquery.min.js"></script> 
        <script type="text/javascript" src="../bootstrap/js/bootstrap.min.js"></script> 
        <title>Office VBA Password remover</title> 
    </head> 
    <body> 
        <div class="container"> 
            <h1>Office VBA Password remover</h1> 
            <?php if($err){ echo "<div class='alert alert-danger'>$err</div>";} ?> 
            <form method="post" action="index.php" enctype="multipart/form-data" class="form-inline"> 
                <label class="control-label">Office File (doc,docm,xls,xlsm,ppt,pptm): 
                <input type="file" name="excel" class="form-control" required /></label><br /> 
                <input type="submit" class="btn btn-primary" value="Decrypt VBA" /> 
            </form> 
            <h2>How it works</h2> 
            <ol> 
                <li>Upload your Office document. You get a document back</li> 
                <li>Open the downloaded document and press <kbd>ALT + F11</kbd>. Confirm error message about invalid entry "BPx"</li> 
                <li>In the Macro window, <b>do not expand the project</b>, go to Tools > VBA Project Properties</li> 
                <li>On the "Protection" Tab, set a password of your choice <b>and leave the checkbox selected</b>.</li> 
                <li>Save the document and close the Editor</li> 
                <li>Repeat Step 3</li> 
                <li>On the "Protection" Tab, clear the checkbox and password fields</li> 
                <li>Save document again</li> 
                <li>The password is now removed and you can view or change the code as if it was never protected</li> 
            </ol> 
            <hr /> 
            <a href="?source">View Source code</a> 
        </div> 
    </body> 
</html>