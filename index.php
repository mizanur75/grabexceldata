<?php
  session_start(); 
?>
<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Grab Excel Data</title>
    <link href="assets/css/bootstrap.min.css" rel="stylesheet">
  </head>
  <body>
    <div class="container">
      <h1 class="text-center m-5">Grab Excel Data</h1>
      <?php
        if(isset($_SESSION['success'])){
        ?>
          <div class="alert alert-success alert-dismissible fade show" role="alert">
            <?php echo $_SESSION['success']; ?>
            <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
          </div>
       
       <?php 
       }
        unset($_SESSION['success']);
        
      ?>
      <?php
        if(isset($_SESSION['error'])){
        ?>
          <div class="alert alert-warning alert-dismissible fade show" role="alert">
            <?php echo $_SESSION['error']; ?>
            <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
          </div>
       
       <?php 
       }
         unset($_SESSION['error']);
      ?>

      <?php
        if(isset($_SESSION['button'])){
        ?>
          <div class="alert alert-info alert-dismissible fade show" role="alert">
            <?php echo $_SESSION['button']; ?>
            <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
          </div>
       
       <?php 
       }
         unset($_SESSION['button']);
      ?>
      <form action="process.php" method="POST" enctype="multipart/form-data">
        <div class="row">
          <div class="input-group mb-3">
            <input type="file" name="excel_file" class="form-control" id="inputGroupFile02">
            <label class="input-group-text" for="inputGroupFile02">Upload</label>
          </div>
          <input type="submit" name="submit" class="btn btn-success" value="Submit">
        </div>
      </form>

    </div>
    <script src="assets/js/bootstrap.bundle.min.js"></script>
  </body>
</html>