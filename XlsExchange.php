<?php
require_once 'Classes/PHPExcel.php';

class XlsExchange
{
    protected $path_to_input_json_file;
    protected $path_to_output_xlsx_file;
    protected $ftp_host;
    protected $ftp_login;
    protected $ftp_password;
    protected $ftp_dir;
    protected $connection = false;

    /**
     * Установка пути входного json-файла
     * @param string $path_to_input_json_file Путь входного json-файла
     */
    public function setInputFile($path_to_input_json_file): XlsExchange
    {
        $this->path_to_input_json_file = $path_to_input_json_file;
        return $this;
    }

    /**
     * Установка пути выходного xlsx-файла
     * @param string $path_to_output_xlsx_file Путь выходного xlsx-файла
     */
    public function setOutputFile($path_to_output_xlsx_file): XlsExchange
    {
        $this->path_to_output_xlsx_file = $path_to_output_xlsx_file;
        return $this;
    }

    /**
     * Установка данных FTP-сервера
     * @param string $ftp_host Хостинг сервера
     * @param string $ftp_login Имя пользователя FTP
     * @param string $ftp_password Пароль для подключения к FTP
     * @param string $ftp_dir Директория, куда будет сохранен выходной файл
     * @return $this
     */
    public function setFtpData($ftp_host, $ftp_login, $ftp_password, $ftp_dir): XlsExchange
    {
        $this->ftp_host = $ftp_host;
        $this->ftp_login = $ftp_login;
        $this->ftp_password = $ftp_password;
        $this->ftp_dir = $ftp_dir;
        return $this;
    }


    public function parseJson(): array
    {
        $stack = array();
        if (file_exists($this->path_to_input_json_file)){
            $json = file_get_contents($this->path_to_input_json_file);
        }
        else {
            echo "Input json-file not exist";
            exit();
        }
        $obj = json_decode($json, true);
        foreach ($obj['items'] as $item) {
            $id = $item['id'];
            if (strlen((string)$item['item']['barcode']) == 13)
            {
                $barcode = $item['item']['barcode'];
            }
            else {
                $barcode = "Not valid barcode";
            }
            $name = $item['item']['name'];
            $quantity = $item['quantity'];
            $price = $item['price'];
            array_push($stack, [
                    'id' => $id,
                    'barcode' => $barcode,
                    'name' => $name,
                    'quantity' => $quantity,
                    'price' => $price]);
        }
        return $stack;
    }


    /**
     * @param array $stack Массив элементов для записи в xlsx-файл
     */
    public function makeXlsx(array $stack)
    {
        /**
         * @var integer $startLine Начальная строка при записи
         * @var integer $columnPosition  Начальный столбец при записи
         */
        $startLine = 1;
        $columnPosition = 0;

        $document = new PHPExcel(); // Создаем документ
        $sheet = $document->setActiveSheetIndex(0);
        // Устанавливаем ширину столбцов
        $sheet->getColumnDimension('A')->setWidth(11);
        $sheet->getColumnDimension('B')->setWidth(27);
        $sheet->getColumnDimension('C')->setAutoSize(true);
        $sheet->getColumnDimension('D')->setWidth(10);
        $sheet->getColumnDimension('E')->setWidth(10);
        // Устанавливаем числовой формат для столбца со штрихкодами
        $sheet->getStyle('B')->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
        // Устанавливаем заголовки столбцов
        $titles = array("ID", "ШК", "Название", "Количество", "Сумма");
        foreach ($titles as $title) {
            $sheet->setCellValueByColumnAndRow($currentColumn, $startLine, $title);
            $currentColumn++;
        }
        // Записываем данные о товарах
        foreach ($stack as $key => $item) {
            // Перекидываем указатель на следующую строку
            $startLine++;
            // Указатель на первый столбец
            $currentColumn = $columnPosition;
            foreach ($item as $value) {
                $sheet->setCellValueByColumnAndRow($currentColumn, $startLine, $value);
                $currentColumn++;
            }
        }
        // Сохраняем файл
        try {
            $objWriter = \PHPExcel_IOFactory::createWriter($document, 'Excel2007');
            $objWriter->save($this->path_to_output_xlsx_file);
            echo "File was saved on local-server\n";
        } catch (PHPExcel_Writer_Exception $e) {
            echo $e;
            exit();
        }
    }
    public function uploadToFTP()
    {
        $conn_id = ftp_connect($this->ftp_host) or die("No connection");
        $login_result = ftp_login($conn_id, $this->ftp_login, $this->ftp_password);
        if ((!$conn_id) || (!$login_result)) {
            echo "FTP connection has failed!";
            echo "Attempted to connect to $this->ftp_host for user: $this->ftp_login";
            exit;
        } else {
            echo "Connected to $this->ftp_host, for user: $this->ftp_login";
            $this->connection = true;
        }
        $dirs = ftp_nlist($conn_id, "");
        $path = "tmp";
        if (!in_array($path, $dirs)) {
            ftp_mkdir($conn_id, $path);
        }
        $upload = ftp_put($conn_id, $this->ftp_dir, $this->path_to_output_xlsx_file, FTP_ASCII);
        if (!upload)
        {
            echo "Upload failed";
            exit;
        }
        else
        {
            echo "Upload successful";
        }
        ftp_close($conn_id);
    }

    public function export()
    {
        $stack = $this->parseJson();
        $this->makeXlsx($stack);
        if ($this->ftp_host && $this->ftp_login && $this->ftp_password && $this->ftp_dir){
            $this->uploadToFTP();
        }

    }
}



