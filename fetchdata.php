<?php
if(empty($_SERVER['HTTP_X_REQUESTED_WITH']) || strtolower($_SERVER['HTTP_X_REQUESTED_WITH']) !== 'xmlhttprequest') {
    die('Direct access not permitted');
}
//Vertify ajax
if (!isset($_POST['nonce']) || $_POST['nonce'] !== md5('export-feedback-from-mail')) {
    die('-1');
}

$list_plugin = array(
    'wp-speed-of-light' => 'WP Speed Of Light',
    'wp-meta-seo' => 'WP Meta Seo',
    'wp-latest-posts' => 'WP Latest Posts',
    'imagerecycle-pdf-image-compression' => 'ImageRecycle pdf & image compression',
    'advanced-gutenberg' => 'Advanced Gutenberg'
);

if (isset($_POST['plugin'])) {
    $datas = get_datas_from_mail($list_plugin[$_POST['plugin']]);
    echo json_encode($datas);
    die;
}

function get_datas_from_mail($selected_plugin) {
    if (!function_exists('imap_open')) {
        return array('status' => false, 'message' => 'Check IMAP module configuration!');
    }
    //Email account
    $email = array(
        'host' => '{imap.joomunited.com:993/imap/ssl/novalidate-cert}INBOX',
        'username' => 'feedback@joomunited.com',
        'password' => 'aW9kupei5ru?yie'
    );

    $connection = imap_open($email['host'], $email['username'], $email['password']);
    if (!$connection) {
        return array('status' => false, 'message' => imap_last_error());
    }
    /* Search Emails having the specified keyword in the email subject */
    $emailData = imap_search($connection, 'SUBJECT "Feedback from" BODY "plugin feedback from" BODY "Feedback Information" BODY "Wordpress version"');

    if (empty($emailData)) {
        return array('status' => false, 'message' => 'Could not find feedback email!');
    }
    arsort($emailData);
    // Filter list message with selected plugin
    $needed = array();
    foreach ($emailData as $email) {
        //Fetch body of the email.
        $message = imap_body($connection, $email, 2);
        if (preg_match('#<i>(.*?)</i>#Usmi', $message, $matches)) {
            if (!empty($matches[1]) && $selected_plugin === $matches[1]) {
                $needed[] = $message;
            }
        }
    }

    imap_close($connection);
    // Not found message body
    if (empty($needed)) {
        return array('status' => false, 'message' => 'No messages found for the selected plugin');
    }

    $result = array();
    foreach ($needed as $need) {
        if (preg_match('#<pre>(.*?)</pre>#Usmi', $need, $matches)) {
            if (!empty($matches[1])) {
                $result[] = json_decode($matches[1], true);
            }
        }
    }

    if (!empty($result)) {
        return array('status' => true, 'message' => $result);
    }

    return array('status' => false, 'message' => 'Not found feedback data yet');
}