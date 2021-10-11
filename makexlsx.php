<?php
if(empty($_SERVER['HTTP_X_REQUESTED_WITH']) || strtolower($_SERVER['HTTP_X_REQUESTED_WITH']) !== 'xmlhttprequest') {
    die('Direct access not permitted');
}
//Vertify ajax
if (!isset($_POST['nonce']) || $_POST['nonce'] !== md5('export-feedback-from-mail')) {
    die('-1');
}

$datas = make_table_for_datas($_POST['datas']);
$file = generate_excel_file($datas);
echo json_encode($file); die;

//Create array for making table
function make_table_for_datas($datas) {
    $default_result = array(
        array('', 'Features', '', '', '', '', '', '', '', '', '', 'General informations', '', '', '', 'Theme', '', '', 'Server', '', '', '', '', '', '', '', '', '', 'Database', '', '', 'Wordpress constants', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'Filesystem permissions', '', '', '', '', '', '', 'Media handling', '', '', '', '', 'Installed plugins'),
        array('Timestamp', 'Feature missing', 'Feature missing comment', 'Not working', 'Not working comment', 'Found better', 'Found better comment', 'Searching something else', 'Searching something else comment', 'Other', 'Other comment', 'Share infos', 'WP Version', 'Https', 'Multisite', 'Active theme', 'Active Theme version', 'Active Theme folder', 'Web server', 'PHP version', 'PHP SAPI', 'PHP max input variables', 'PHP time limit', 'PHP memory limit', 'Max input time', 'Upload max filesize', 'PHP post max size', 'cURL version', 'Extension', 'Version', 'Client', 'ABSPATH', 'WP_HOME', 'WP_SITEURL', 'WP_CONTENT_DIR', 'WP_PLUGIN_DIR', 'WP_MAX_MEMORY_LIMIT', 'WP_DEBUG', 'WP_DEBUG_DISPLAY', 'WP_DEBUG_LOG', 'SCRIPT_DEBUG', 'WP_CACHE', 'CONCATENATE_SCRIPTS', 'COMPRESS_SCRIPTS', 'COMPRESS_CSS', 'WP_LOCAL_DEV', 'The main WordPress directory', 'The wp-content directory', 'The uploads directory', 'The plugins directory', 'The themes directory', 'The must use plugins directory', 'WordPress configuration file', 'Active editor', 'ImageMagick version number', 'ImageMagick version string', 'GD version', 'Ghostscript version'),
    );

    $list_install_plugin = array();
    $items = array();

    foreach ($datas as $data) {
        //
        // Filter list installed plugin
        if (!empty($data['Active plugins'])) {
            foreach ($data['Active plugins'] as $v) {
                $list_install_plugin[] = $v['Name'];
            }
        }
        // FIlter timestamp
        $timestamp = (isset($data['Timestamp']) && $data['Timestamp']) ? $data['Timestamp'] : '';
        //
        // Filter reasons
        $feature_missing = 'NO';
        $feature_missing_comment = '';
        $not_working = 'NO';
        $not_working_comment = '';
        $found_better = 'NO';
        $found_better_comment = '';
        $searh_something = 'NO';
        $searh_something_comment = '';
        $other = 'NO';
        $other_comment = '';

        if (!empty($data['Reasons'])) {
            foreach ($data['Reasons'] as $v) {
                if ($v['reason'] === 'There\'s a feature missing') {
                    $feature_missing = 'YES';
                    $feature_missing_comment = $v['comment'];
                }
                if ($v['reason'] === 'The plugin is not working great') {
                    $not_working = 'YES';
                    $not_working_comment = $v['comment'];
                }
                if ($v['reason'] === 'I found a better plugin') {
                    $found_better = 'YES';
                    $found_better_comment = $v['comment'];
                }
                if ($v['reason'] === 'I was searching for something else') {
                    $searh_something = 'YES';
                    $searh_something_comment = $v['comment'];
                }
                if ($v['reason'] === 'Other, We\'d like to hear your opinion :)') {
                    $other = 'YES';
                    $other_comment = $v['comment'];
                }
            }
        }
        //
        // Filter General infomations
        $share_infos = (isset($data['Wordpress version']) && $data['Wordpress version'])? 'YES' : 'NO';
        $wp_version = (isset($data['Wordpress version']) && $data['Wordpress version'])  ? $data['Wordpress version'] : '';
        $is_https = (isset($data['Is https site']) && $data['Is https site']) ? $data['Is https site'] : 'NO';
        $is_multisite = (isset($data['Is multisite']) && $data['Is multisite']) ? $data['Is multisite'] : 'NO';
        //
        // Filter active theme
        $active_theme = (isset($data['Active theme']['Name']) && $data['Active theme']['Name']) ? $data['Active theme']['Name'] : '';
        $active_theme_version = (isset($data['Active theme']['Version']) && $data['Active theme']['Version']) ? $data['Active theme']['Version'] : '';
        $active_theme_folder = (isset($data['Active theme']['Folder']) && $data['Active theme']['Folder']) ? $data['Active theme']['Folder'] : '';
        //
        // Filter server info
        $web_server = (isset($data['Server']['Web server']) && $data['Server']['Web server']) ? $data['Server']['Web server'] : '';
        $php_version = (isset($data['Server']['PHP version']) && $data['Server']['PHP version']) ? $data['Server']['PHP version'] : '';
        $php_sapi = (isset($data['Server']['PHP SAPI']) && $data['Server']['PHP SAPI']) ? $data['Server']['PHP SAPI'] : '';
        $php_max_input = (isset($data['Server']['PHP max input variables']) && $data['Server']['PHP max input variables']) ? $data['Server']['PHP max input variables'] : '';
        $php_time_limit = (isset($data['Server']['PHP time limit']) && $data['Server']['PHP time limit']) ? $data['Server']['PHP time limit'] : '';
        $php_memory_limit = (isset($data['Server']['PHP memory limit']) && $data['Server']['PHP memory limit']) ? $data['Server']['PHP memory limit'] : '';
        $max_input_limit = (isset($data['Server']['Max input time']) && $data['Server']['Max input time']) ? $data['Server']['Max input time'] : '';
        $upload_max_size = (isset($data['Server']['Upload max filesize']) && $data['Server']['Upload max filesize']) ? $data['Server']['Upload max filesize'] : '';
        $post_max_size = (isset($data['Server']['PHP post max size'] ) && $data['Server']['PHP post max size'] ) ? $data['Server']['PHP post max size'] : '';
        $curl_version = (isset($data['Server']['cURL version']) && $data['Server']['cURL version']) ? $data['Server']['cURL version'] : '';
        //
        // Filter database
        $db_extension = (isset($data['Database']['Extension']) && $data['Database']['Extension']) ? $data['Database']['Extension'] : '';
        $db_version = (isset($data['Database']['Version']) && $data['Database']['Version']) ? $data['Database']['Version'] : '';
        $db_client = (isset($data['Database']['Client']) && $data['Database']['Client']) ? $data['Database']['Client'] : '';

        //
        // Filter WP constants
        $abspath = (isset($data['Wordpress constants']['ABSPATH']) && $data['Wordpress constants']['ABSPATH']) ? $data['Wordpress constants']['ABSPATH'] : '';
        $wp_home = (isset($data['Wordpress constants']['WP_HOME']) && $data['Wordpress constants']['WP_HOME']) ? $data['Wordpress constants']['WP_HOME'] : '';
        $wp_siteurl = (isset($data['Wordpress constants']['WP_SITEURL']) && $data['Wordpress constants']['WP_SITEURL']) ? $data['Wordpress constants']['WP_SITEURL'] : '';
        $wp_content_dir = (isset($data['Wordpress constants']['WP_CONTENT_DIR']) && $data['Wordpress constants']['WP_CONTENT_DIR']) ? $data['Wordpress constants']['WP_CONTENT_DIR'] : '';
        $wp_plugin_dir = (isset($data['Wordpress constants']['WP_PLUGIN_DIR']) && $data['Wordpress constants']['WP_PLUGIN_DIR']) ? $data['Wordpress constants']['WP_PLUGIN_DIR'] : '';
        $wp_max_memory_limit = (isset($data['Wordpress constants']['WP_MAX_MEMORY_LIMIT']) && $data['Wordpress constants']['WP_MAX_MEMORY_LIMIT']) ? $data['Wordpress constants']['WP_MAX_MEMORY_LIMIT'] : '';
        $wp_debug = (isset($data['Wordpress constants']['WP_DEBUG']) && $data['Wordpress constants']['WP_DEBUG']) ? $data['Wordpress constants']['WP_DEBUG'] : '';
        $wp_debug_display = (isset($data['Wordpress constants']['WP_DEBUG_DISPLAY']) && $data['Wordpress constants']['WP_DEBUG_DISPLAY']) ? $data['Wordpress constants']['WP_DEBUG_DISPLAY'] : '';
        $wp_debug_log = (isset($data['Wordpress constants']['WP_DEBUG_LOG']) && $data['Wordpress constants']['WP_DEBUG_LOG']) ? $data['Wordpress constants']['WP_DEBUG_LOG'] : '';
        $script_debug = (isset($data['Wordpress constants']['SCRIPT_DEBUG']) && $data['Wordpress constants']['SCRIPT_DEBUG']) ? $data['Wordpress constants']['SCRIPT_DEBUG'] : '';
        $wp_cache = (isset($data['Wordpress constants']['WP_CACHE']) && $data['Wordpress constants']['WP_CACHE']) ? $data['Wordpress constants']['WP_CACHE'] : '';
        $concatenate_scripts = (isset($data['Wordpress constants']['CONCATENATE_SCRIPTS']) && $data['Wordpress constants']['CONCATENATE_SCRIPTS']) ? $data['Wordpress constants']['CONCATENATE_SCRIPTS'] : '';
        $compress_scripts = (isset($data['Wordpress constants']['COMPRESS_SCRIPTS']) && $data['Wordpress constants']['COMPRESS_SCRIPTS']) ? $data['Wordpress constants']['COMPRESS_SCRIPTS'] : '';
        $compress_css = (isset($data['Wordpress constants']['COMPRESS_CSS']) && $data['Wordpress constants']['COMPRESS_CSS']) ? $data['Wordpress constants']['COMPRESS_CSS'] : '';
        $wp_local_dev = (isset($data['Wordpress constants']['WP_LOCAL_DEV']) && $data['Wordpress constants']['WP_LOCAL_DEV']) ? $data['Wordpress constants']['WP_LOCAL_DEV'] : '';
        // Filter file permissions
        $file_main_wp_dir = (isset($data['Filesystem permissions']['The main WordPress directory'])  && $data['Filesystem permissions']['The main WordPress directory']) ? $data['Filesystem permissions']['The main WordPress directory'] : '';
        $file_wp_content_dir = (isset($data['Filesystem permissions']['The wp-content directory']) && $data['Filesystem permissions']['The wp-content directory']) ? $data['Filesystem permissions']['The wp-content directory'] : '';
        $file_up_dir = (isset($data['Filesystem permissions']['The uploads directory']) && $data['Filesystem permissions']['The uploads directory']) ? $data['Filesystem permissions']['The uploads directory'] : '';
        $file_plugin_dir = (isset($data['Filesystem permissions']['The plugins directory']) && $data['Filesystem permissions']['The plugins directory']) ? $data['Filesystem permissions']['The plugins directory'] : '';
        $file_theme_dir = (isset($data['Filesystem permissions']['The themes directory']) && $data['Filesystem permissions']['The themes directory']) ? $data['Filesystem permissions']['The themes directory'] : '';
        $file_must_use_plugin_dir = (isset($data['Filesystem permissions']['The must use plugins directory']) && $data['Filesystem permissions']['The must use plugins directory']) ? $data['Filesystem permissions']['The must use plugins directory'] : '';
        $file_wp_configuration_file = (isset($data['Filesystem permissions']['WordPress configuration file']) && $data['Filesystem permissions']['WordPress configuration file']) ? $data['Filesystem permissions']['WordPress configuration file'] : '';
        //
        // Filter media handle
        $active_editor = (isset($data['Media handling']['Active editor']) && $data['Media handling']['Active editor']) ? $data['Media handling']['Active editor'] : '';
        $image_version_number = (isset($data['Media handling']['ImageMagick version number']) && $data['Media handling']['ImageMagick version number']) ? $data['Media handling']['ImageMagick version number'] : '';
        $image_version_string = (isset($data['Media handling']['ImageMagick version string']) && $data['Media handling']['ImageMagick version string']) ? $data['Media handling']['ImageMagick version string'] : '';
        $gd_version = (isset($data['Media handling']['GD version']) && $data['Media handling']['GD version']) ? $data['Media handling']['GD version'] : '';
        $ghostscript_version = (isset($data['Media handling']['Ghostscript version']) && $data['Media handling']['Ghostscript version']) ? $data['Media handling']['Ghostscript version'] : '';


        $items[] = array(
            $timestamp,
            $feature_missing,
            $feature_missing_comment,
            $not_working,
            $not_working_comment,
            $found_better,
            $found_better_comment,
            $searh_something,
            $searh_something_comment,
            $other,
            $other_comment,
            $share_infos,
            $wp_version,
            $is_https,
            $is_multisite,
            $active_theme,
            $active_theme_version,
            $active_theme_folder,
            $web_server,
            $php_version,
            $php_sapi,
            $php_max_input,
            $php_time_limit,
            $php_memory_limit,
            $max_input_limit,
            $upload_max_size,
            $post_max_size,
            $curl_version,
            $db_extension,
            $db_version,
            $db_client,
            $abspath,
            $wp_home,
            $wp_siteurl,
            $wp_content_dir,
            $wp_plugin_dir,
            $wp_max_memory_limit,
            $wp_debug,
            $wp_debug_display,
            $wp_debug_log,
            $script_debug,
            $wp_cache,
            $concatenate_scripts,
            $compress_scripts,
            $compress_css,
            $wp_local_dev,
            $file_main_wp_dir,
            $file_wp_content_dir,
            $file_up_dir,
            $file_plugin_dir,
            $file_theme_dir,
            $file_must_use_plugin_dir,
            $file_wp_configuration_file,
            $active_editor,
            $image_version_number,
            $image_version_string,
            $gd_version,
            $ghostscript_version
            );
    }

    $list_install_plugin = array_unique($list_install_plugin);
    sort($list_install_plugin);
    // Make row2
    $default_result[1] = array_merge($default_result[1], $list_install_plugin);

    // Get value of of installed plugin
    $value_installed_plugin = array();
    if (!empty($datas)) {
        foreach ($datas as $data) {
            $reverse_installed_plugin = array_fill_keys($list_install_plugin, '');

            if (!empty($data['Active plugins'])) {
                foreach ($data['Active plugins'] as $v) {

                    if (array_key_exists($v['Name'], $reverse_installed_plugin)) {
                        $reverse_installed_plugin[$v['Name']] = get_version_number($v['Version']);
                    }
                }
            }
            $value_installed_plugin[] = $reverse_installed_plugin;
        }
    }

    if (!empty($items)) {
        // Make list plugin value of child
        foreach ($items as $k => $item) {
            $items[$k] = array_merge($item, array_values($value_installed_plugin[$k]));
        }
    }

    //Create child
    // Make row3,4,5..
    if (!empty($items)) {
        foreach ($items as $item) {
            array_push($default_result, $item);
        }
    }

    return $default_result;
}



function generate_excel_file($datas) {
    if (!is_array($datas)) {
        return array('status' => false, 'file' => 'Unknown data!');
    }

    require_once 'phpspreadsheet/vendor/autoload.php';

    $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
    $activeSheet = $spreadsheet->getActiveSheet();

    $activeSheet->fromArray($datas);
    $maxCell = $activeSheet->getHighestRowAndColumn();
    $activeSheet->rangeToArray('A1:' . $maxCell['column'] . $maxCell['row'], '', false);
    $activeSheet->mergeCells('B1:K1');
    $activeSheet->mergeCells('L1:O1');
    $activeSheet->mergeCells('P1:R1');
    $activeSheet->mergeCells('S1:AB1');
    $activeSheet->mergeCells('AC1:AE1');
    $activeSheet->mergeCells('AF1:AT1');
    $activeSheet->mergeCells('AU1:BA1');
    $activeSheet->mergeCells('BB1:BF1');
    $activeSheet->mergeCells('BG1:'.$maxCell['column'].'1');
    $objWriter = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);

    ob_start();
    $objWriter->save("php://output");
    $xlsxData = ob_get_contents();
    ob_end_clean();

    return array(
        'status' => true,
        'file' => "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,".base64_encode($xlsxData)
    );
}


function get_version_number($strings) {
    if (strpos($strings, '{{version}}') !== false) {
        return '{{version}}';
    }

    if (preg_match('/(\d+\.?)+/', $strings, $matches)) {
        return $matches[0];
    }

    return 'undefined';
}