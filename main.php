<?php
    require_once 'XlsExchange.php';
    (new \XlsExchange())
        ->setInputFile('tmp/order.json')
        ->setOutputFile('tmp/items.xlsx')
        ->setFtpData('localhost', 'ftp','ftp','/tmp/items.xlsx')
        ->export();
