<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Export Feedback</title>
    <!-- Custom style/script -->
    <link rel="stylesheet" type="text/css" href="assets/css/export.css">
    <script type="text/javascript" src="assets/js/jquery.min.js"></script>
</head>
<body>
<div id="export-mail" class="container">
    <div id="introduction" class="header">
        <h1>Export plugin feedback</h1>
    </div>
    <div id="content">
        <div class="loading process">
            <label class="export-label">Status</label>
            <span class="status">Ready!</span>
        </div>
        <form action="">
            <input type="hidden" id="nonce" class="nonce" value="<?php echo md5('export-feedback-from-mail')?>" />
            <label class="export-label">Select plugin</label>
            <select name="plugin" class="select-plugin">
                <option value="wp-speed-of-light">WP Speed Of Light</option>
                <option value="wp-meta-seo">WP Meta SEO</option>
                <option value="wp-latest-posts">WP Latest Posts</option>
                <option value="imagerecycle-pdf-image-compression">ImageRecycle pdf & image compression</option>
                <option value="advanced-gutenberg">Advanced Gutenberg</option>
            </select>
            <div class="clear"></div>
            <input type="button" value="Export data" id="export-data" class="export-data" />
        </form>
        <div class="clear"></div>
    </div>
    <script type="text/javascript" src="assets/js/export.js"></script>
</div>
</body>
</html>