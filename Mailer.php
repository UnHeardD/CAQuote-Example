
<?php



class CAQuote_Mailer

{

    var $message;

    var $FILES;

    var $EMAIL;



    public function CAQuote_Mailer($to_address, $from_address, $subject, $reply_address=null, $mailer=null, $custom_header=null)

    {

        $this->EMAIL = array(

            "to" => $to_address,

            "from" => $from_address,

            "subject" => $subject,

            "reply" => (empty($reply_address) ? $from_address : $reply_address),

            "mailer" => (empty($mailer) ? "X-Mailer: PHP/" . phpversion() : $mailer),

            "header" => (empty($custom_header) ? "" : $custom_header),

            "boundary" => "_mimeboundary_".md5(uniqid(mt_rand(), 1))

            );



        $this->message = "";



        $this->FILES = array();

    }



    public function addFile($filename, $type=null, $filecontents=null)

    {

        if ($filecontents !== null)

        {

            $index = count($this->FILES);

            $this->FILES[$index]['data'] = chunk_split(base64_encode($filecontents));

            $this->FILES[$index]['name'] = basename($filename);



            if (empty($type))

                $this->FILES[$index]['mime'] = mime_content_type($filename);

            else

                $this->FILES[$index]['mime'] = $type;

        }

        else if (file_exists($filename))

        {

            $index = count($this->FILES);

            $this->FILES[$index]['data'] = chunk_split(base64_encode(file_get_contents($filename)));

            $this->FILES[$index]['name'] = basename($filename);



            if (empty($type))

                $this->FILES[$index]['mime'] = mime_content_type($filename);

            else

                $this->FILES[$index]['mime'] = $type;

        }

        else

        {

            $this->Error_Handle("File specified -- {$filename} -- does not exist.");

        }

    }







    public function addText($text)

    {

        $this->message .= $text;

    }





    public function getHeader()

    {

        $header = "From: {$this->EMAIL['from']}\r\n"

                . "Reply-To: {$this->EMAIL['reply']}\r\n"

                . "X-Mailer: {$this->EMAIL['mailer']}\r\n"

                . "MIME-Version: 1.0\r\n"

                . "Content-Type: multipart/mixed; boundary=\"{$this->EMAIL['boundary']}\";\r\n";



        return $header;

    }

    

    public function getEmail()

    {    

        $content .= "--{$this->EMAIL['boundary']}\r\n"

                . "Content-Type: text/plain; charset=\"iso-8859-1\"\r\n"

                . "Content-Transfer-Encoding: 7bit\r\n\r\n"

                . $this->message . "\r\n";



        if (!empty($this->FILES))

        {

            foreach($this->FILES as $file)

            {

                $content .= "--{$this->EMAIL['boundary']}\r\n"

                . "Content-Type: {$file['mime']}; name=\"{$file['name']}\"\r\n"

                . "Content-Transfer-Encoding: base64\r\n"

                . "Content-Disposition: attachment\r\n\r\n"

                . $file['data'] . "\r\n";

            }

        }



        $content .= "--{$this->EMAIL['boundary']}--\r\n";





        return $content;

    }



    public function send()

    {

        $result = mail($this->EMAIL['to'], $this->EMAIL['subject'], $this->getEmail(), $this->getHeader());


    }





    public function Error_Handle($error)

    {

        die($error);

    }

}
?>