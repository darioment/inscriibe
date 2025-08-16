<?php
error_reporting(E_ALL);
ini_set('display_errors', 1);

require 'vendor/autoload.php';

if(isset($_POST["nombre"])){

  $nombre = 'inscripcion_'.$_POST["apellido1"]." ".$_POST["apellido2"]." ".$_POST["nombre"].'.docx';

  $fechanac = $_POST["fechanac"];
  $dateParts = explode("/", $fechanac);

  if (count($dateParts) === 3) {
      $dianac = $dateParts[1];
      $mesnac = $dateParts[0];
      $anionac = substr($dateParts[2], -2); // Extraer los últimos dos dígitos del año

      $fechaActual = date("m/d/Y"); // Obtenemos la fecha actual en el mismo formato
      
      $fechaNacimientoObj = DateTime::createFromFormat("m/d/Y", $fechanac);
      $fechaActualObj = DateTime::createFromFormat("m/d/Y", $fechaActual);
      
      $diff = $fechaNacimientoObj->diff($fechaActualObj);
      
      $edad = $diff->y;
      $edadmes = $diff->m;      
  } else {
      echo "Formato de fecha incorrecto.";
  }

  $fechanacpa = $_POST["fechanacpa"];
  $dateParts = explode("/", $fechanacpa);

  if (count($dateParts) === 3) {
      $dianacpa = $dateParts[1];
      $mesnacpa = $dateParts[0];
      $anionacpa = substr($dateParts[2], -2); // Extraer los últimos dos dígitos del año

      $fechaActualpa = date("m/d/Y"); // Obtenemos la fecha actual en el mismo formato
      
      $fechaNacimientoObjpa = DateTime::createFromFormat("m/d/Y", $fechanacpa);
      $fechaActualObjpa = DateTime::createFromFormat("m/d/Y", $fechaActualpa);
      
      $diffpa = $fechaNacimientoObjpa->diff($fechaActualObjpa);
      
      $edadpa = $diffpa->y;
      $edadmespa = $diffpa->m;      
  } else {
      echo "Formato de fecha tutor incorrecto.";
  }

  // Procesamiento de fecha para segundo tutor
  $fechanacpa2 = $_POST["fechanacpa2"];
  $dateParts2 = explode("/", $fechanacpa2);

  if (count($dateParts2) === 3) {
      $dianacpa2 = $dateParts2[1];
      $mesnacpa2 = $dateParts2[0];
      $anionacpa2 = substr($dateParts2[2], -2); // Extraer los últimos dos dígitos del año

      $fechaActualpa2 = date("m/d/Y"); // Obtenemos la fecha actual en el mismo formato
      
      $fechaNacimientoObjpa2 = DateTime::createFromFormat("m/d/Y", $fechanacpa2);
      $fechaActualObjpa2 = DateTime::createFromFormat("m/d/Y", $fechaActualpa2);
      
      $diffpa2 = $fechaNacimientoObjpa2->diff($fechaActualObjpa2);
      
      $edadpa2 = $diffpa2->y;
      $edadmespa2 = $diffpa2->m;      
  } else {
      echo "Formato de fecha segundo tutor incorrecto.";
  }

  if(!file_exists('docs_himno/'.$nombre)){
    $templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor('inscribe.docx');
    $templateProcessor->setValue('apellido1', $_POST["apellido1"]);
    $templateProcessor->setValue('apellido2',$_POST["apellido2"]);
    $templateProcessor->setValue('nombre', $_POST["nombre"]);
    $templateProcessor->setValue('dia',$dianac);
    $templateProcessor->setValue('mes',$mesnac);
    $templateProcessor->setValue('anio',$anionac);
    $templateProcessor->setValue('edad',$edad);
    $templateProcessor->setValue('edadmes',$edadmes);
    $templateProcessor->setValue('sexo',$_POST["sexo"]);
    $templateProcessor->setValue('curp',$_POST["curp"]);
    $templateProcessor->setValue('peso',$_POST["peso"]);
    $templateProcessor->setValue('talla',$_POST["talla"]);
    $templateProcessor->setValue('lentes',$_POST["lentes"]);
    $templateProcessor->setValue('zapatos',$_POST["zapatos"]);
    $templateProcessor->setValue('cartilla',$_POST["cartilla"]);
    $templateProcessor->setValue('vacunas',$_POST["vacunas"]);
    $templateProcessor->setValue('nacionalidad',$_POST["nacionalidad"]);
    $templateProcessor->setValue('entidadnac',$_POST["entidadnac"]);
    $templateProcessor->setValue('serviciomed',$_POST["serviciomed"]);
    $templateProcessor->setValue('calle',$_POST["calle"]);
    $templateProcessor->setValue('entrecalle1',$_POST["entrecalle1"]);
    $templateProcessor->setValue('entrecalle2',$_POST["entrecalle2"]);
    $templateProcessor->setValue('numext',$_POST["numext"]);
    $templateProcessor->setValue('numint',$_POST["numint"]);
    $templateProcessor->setValue('serviciomed',$_POST["serviciomed"]);
    $templateProcessor->setValue('manzana',$_POST["manzana"]);
    $templateProcessor->setValue('lote',$_POST["lote"]);
    $templateProcessor->setValue('depto',$_POST["depto"]);
    $templateProcessor->setValue('colonia',$_POST["colonia"]);
    $templateProcessor->setValue('cp',$_POST["cp"]);
    $templateProcessor->setValue('localidad',$_POST["localidad"]);
    $templateProcessor->setValue('municipio',$_POST["municipio"]);
    $templateProcessor->setValue('referencia',$_POST["referencia"]);
    $templateProcessor->setValue('entidad',$_POST["entidad"]);
    $templateProcessor->setValue('telcasa',$_POST["telcasa"]);
    $templateProcessor->setValue('telcel',$_POST["telcel"]);
    $templateProcessor->setValue('correo1',$_POST["correo1"]);
    $templateProcessor->setValue('correo2',$_POST["correo2"]);
    $templateProcessor->setValue('hermanos',$_POST["hermanos"]);
    $templateProcessor->setValue('redsocial',$_POST["redsocial"]);
    $templateProcessor->setValue('tipored',$_POST["tipored"]);
    $templateProcessor->setValue('correo1',$_POST["correo1"]);
    $templateProcessor->setValue('hermanos',$_POST["hermanos"]);

    $templateProcessor->setValue('grado',$_POST["grado"]);
    $templateProcessor->setValue('grupo',$_POST["grupo"]);
    $templateProcessor->setValue('gradoherm',$_POST["gradoherm"]);
    $templateProcessor->setValue('grupoherm',$_POST["grupoherm"]);

    $templateProcessor->setValue('parentesco',$_POST["parentesco"]);
    $templateProcessor->setValue('apellido1herm',$_POST["apellido1herm"]);
    $templateProcessor->setValue('apellido2herm',$_POST["apellido2herm"]);
    $templateProcessor->setValue('nombreherm',$_POST["nombreherm"]);

    $templateProcessor->setValue('diapa',$dianacpa);
    $templateProcessor->setValue('mespa',$mesnacpa);
    $templateProcessor->setValue('aniopa',$anionacpa);
    $templateProcessor->setValue('edadpa',$edadpa);
    $templateProcessor->setValue('edadmespa',$edadmespa);
    $templateProcessor->setValue('sexopa',$_POST["sexopa"]);
    $templateProcessor->setValue('curppa',$_POST["curppa"]);
    $templateProcessor->setValue('civil',$_POST["civil"]);
    $templateProcessor->setValue('estudios',$_POST["estudios"]);
    $templateProcessor->setValue('laboral',$_POST["laboral"]);
    $templateProcessor->setValue('nacionalidadpa',$_POST["nacionalidadpa"]);
    $templateProcessor->setValue('entidadnacpa',$_POST["entidadnacpa"]);
    $templateProcessor->setValue('documento',$_POST["documento"]);
    $templateProcessor->setValue('callepa',$_POST["callepa"]);
    $templateProcessor->setValue('calle1pa',$_POST["calle1pa"]);
    $templateProcessor->setValue('calle2pa',$_POST["calle2pa"]);
    $templateProcessor->setValue('numextpa',$_POST["numextpa"]);
    $templateProcessor->setValue('numintpa',$_POST["numintpa"]);
    $templateProcessor->setValue('manzanapa',$_POST["manzanapa"]);
    $templateProcessor->setValue('lotepa',$_POST["lotepa"]);
    $templateProcessor->setValue('deptopa',$_POST["deptopa"]);
    $templateProcessor->setValue('coloniapa',$_POST["coloniapa"]);
    $templateProcessor->setValue('cppa',$_POST["cppa"]);
    $templateProcessor->setValue('localidadpa',$_POST["localidadpa"]);
    $templateProcessor->setValue('municipiopa',$_POST["municipiopa"]);
    $templateProcessor->setValue('referenciapa',$_POST["referenciapa"]);
    $templateProcessor->setValue('entidadpa',$_POST["entidadpa"]);
    $templateProcessor->setValue('telcasapa',$_POST["telcasapa"]);
    $templateProcessor->setValue('telcelpa',$_POST["telcelpa"]);
    $templateProcessor->setValue('correo1pa',$_POST["correo1pa"]);
    $templateProcessor->setValue('correo2pa',$_POST["correo2pa"]);
    $templateProcessor->setValue('apellido1pa', $_POST["apellido1pa"]);
    $templateProcessor->setValue('apellido2pa',$_POST["apellido2pa"]);
    $templateProcessor->setValue('nombrepa', $_POST["nombrepa"]);

    // Variables de plantilla para el segundo tutor
    $templateProcessor->setValue('parentescopa2',$_POST["parentescopa2"]);
    $templateProcessor->setValue('diapa2',$dianacpa2);
    $templateProcessor->setValue('mespa2',$mesnacpa2);
    $templateProcessor->setValue('aniopa2',$anionacpa2);
    $templateProcessor->setValue('edadpa2',$edadpa2);
    $templateProcessor->setValue('edadmespa2',$edadmespa2);
    $templateProcessor->setValue('sexopa2',$_POST["sexopa2"]);
    $templateProcessor->setValue('curppa2',$_POST["curppa2"]);
    $templateProcessor->setValue('civilpa2',$_POST["civilpa2"]);
    $templateProcessor->setValue('estudiospa2',$_POST["estudiospa2"]);
    $templateProcessor->setValue('laboralpa2',$_POST["laboralpa2"]);
    $templateProcessor->setValue('nacionalidadpa2',$_POST["nacionalidadpa2"]);
    $templateProcessor->setValue('entidadnacpa2',$_POST["entidadnacpa2"]);
    $templateProcessor->setValue('documentopa2',$_POST["documentopa2"]);
    $templateProcessor->setValue('callepa2',$_POST["callepa2"]);
    $templateProcessor->setValue('calle1pa2',$_POST["calle1pa2"]);
    $templateProcessor->setValue('calle2pa2',$_POST["calle2pa2"]);
    $templateProcessor->setValue('numextpa2',$_POST["numextpa2"]);
    $templateProcessor->setValue('numintpa2',$_POST["numintpa2"]);
    $templateProcessor->setValue('manzanapa2',$_POST["manzanapa2"]);
    $templateProcessor->setValue('lotepa2',$_POST["lotepa2"]);
    $templateProcessor->setValue('deptopa2',$_POST["deptopa2"]);
    $templateProcessor->setValue('coloniapa2',$_POST["coloniapa2"]);
    $templateProcessor->setValue('cppa2',$_POST["cppa2"]);
    $templateProcessor->setValue('localidadpa2',$_POST["localidadpa2"]);
    $templateProcessor->setValue('municipiopa2',$_POST["municipiopa2"]);
    $templateProcessor->setValue('referenciapa2',$_POST["referenciapa2"]);
    $templateProcessor->setValue('entidadpa2',$_POST["entidadpa2"]);
    $templateProcessor->setValue('telcasapa2',$_POST["telcasapa2"]);
    $templateProcessor->setValue('telcelpa2',$_POST["telcelpa2"]);
    $templateProcessor->setValue('correo1pa2',$_POST["correo1pa2"]);
    $templateProcessor->setValue('correo2pa2',$_POST["correo2pa2"]);
    $templateProcessor->setValue('apellido1pa2', $_POST["apellido1pa2"]);
    $templateProcessor->setValue('apellido2pa2',$_POST["apellido2pa2"]);
    $templateProcessor->setValue('nombrepa2', $_POST["nombrepa2"]);

    $templateProcessor->setValue('especial', $_POST["especial"]);
    $templateProcessor->setValue('apoyo',$_POST["apoyo"]);
    $templateProcessor->setValue('indigena', $_POST["indigena"]);    

    $templateProcessor->saveAs('docs_himno/'.$nombre);
  }


  // guarda datos
  $servername = "evolution.usoreal.com";
  $username = "mysql";
  $password = "dment25MY!.";
  $dbname = "sec";
  
  // Crear una conexión
  $conn = new mysqli($servername, $username, $password, $dbname);
  
    // Establecer el conjunto de caracteres a Latin-1
    if (!$conn->set_charset("utf8")) {
      die("Error al establecer el conjunto de caracteres: " . $conn->error);
    }
    
  // Verificar la conexión
  if ($conn->connect_error) {
      die("Conexión fallida: " . $conn->connect_error);
  }

  $query = "INSERT INTO ins_himno (apellido1, apellido2, nombre, dia, mes, anio, edad, edadmes, sexo, curp, peso, talla, lentes, zapatos, cartilla, vacunas, nacionalidad, entidadnac, serviciomed, calle, entrecalle1, entrecalle2, numext, numint, manzana, lote, depto, colonia, cp, localidad, municipio, referencia, entidad, telcasa, telcel, correo1, correo2, hermanos, redsocial, tipored, gradoherm, grupoherm, parentesco, apellido1herm, apellido2herm, nombreherm, diapa, mespa, aniopa, edadpa, edadmespa, sexopa, curppa, civil, estudios, laboral, nacionalidadpa, entidadnacpa, documento, callepa, calle1pa, calle2pa, numextpa, numintpa, manzanapa, lotepa, deptopa, coloniapa, cppa, localidadpa, municipiopa, referenciapa, entidadpa, telcasapa, telcelpa, correo1pa, correo2pa, apellido1pa, apellido2pa, nombrepa, especial, apoyo, indigena, grado, grupo, parentescopa2, apellido1pa2, apellido2pa2, nombrepa2, diapa2, mespa2, aniopa2, edadpa2, edadmespa2, sexopa2, curppa2, civilpa2, estudiospa2, laboralpa2, nacionalidadpa2, entidadnacpa2, documentopa2, callepa2, calle1pa2, calle2pa2, numextpa2, numintpa2, manzanapa2, lotepa2, deptopa2, coloniapa2, cppa2, localidadpa2, municipiopa2, referenciapa2, entidadpa2, telcasapa2, telcelpa2, correo1pa2, correo2pa2) 
  VALUES ('".$_POST["apellido1"]."','".$_POST["apellido2"]."','".$_POST["nombre"]."',".$dianac.",".$mesnac.",".$anionac.",".$edad.",".$edadmes.",'".$_POST["sexo"]."','".$_POST["curp"]."','".$_POST["peso"]."','".$_POST["talla"]."','".$_POST["lentes"]."','".$_POST["zapatos"]."','".$_POST["cartilla"]."','".$_POST["vacunas"]."','".$_POST["nacionalidad"]."','".$_POST["entidadnac"]."','".$_POST["serviciomed"]."','".$_POST["calle"]."','".$_POST["entrecalle1"]."','".$_POST["entrecalle2"]."','".$_POST["numext"]."','".$_POST["numint"]."','".$_POST["manzana"]."','".$_POST["lote"]."','".$_POST["depto"]."','".$_POST["colonia"]."','".$_POST["cp"]."','".$_POST["localidad"]."','".$_POST["municipio"]."','".$_POST["referencia"]."','".$_POST["entidad"]."','".$_POST["telcasa"]."','".$_POST["telcel"]."','".$_POST["correo1"]."','".$_POST["correo2"]."','".$_POST["hermanos"]."','".$_POST["redsocial"]."','".$_POST["tipored"]."','".$_POST["gradoherm"]."','".$_POST["grupoherm"]."','".$_POST["parentesco"]."','".$_POST["apellido1herm"]."','".$_POST["apellido2herm"]."','".$_POST["nombreherm"]."',".$dianacpa.",".$mesnacpa.",".$anionacpa.",".$edadpa.",".$edadmespa.",'".$_POST["sexopa"]."','".$_POST["curppa"]."','".$_POST["civil"]."','".$_POST["estudios"]."','".$_POST["laboral"]."','".$_POST["nacionalidadpa"]."','".$_POST["entidadnacpa"]."','".$_POST["documento"]."','".$_POST["callepa"]."','".$_POST["calle1pa"]."','".$_POST["calle2pa"]."','".$_POST["numextpa"]."','".$_POST["numintpa"]."','".$_POST["manzanapa"]."','".$_POST["lotepa"]."','".$_POST["deptopa"]."','".$_POST["coloniapa"]."','".$_POST["cppa"]."','".$_POST["localidadpa"]."','".$_POST["municipiopa"]."','".$_POST["referenciapa"]."','".$_POST["entidadpa"]."','".$_POST["telcasapa"]."','".$_POST["telcelpa"]."','".$_POST["correo1pa"]."','".$_POST["correo2pa"]."','".$_POST["apellido1pa"]."','".$_POST["apellido2pa"]."','".$_POST["nombrepa"]."','".$_POST["especial"]."','".$_POST["apoyo"]."','".$_POST["indigena"]."','".$_POST["grado"]."','".$_POST["grupo"]."','".$_POST["parentescopa2"]."','".$_POST["apellido1pa2"]."','".$_POST["apellido2pa2"]."','".$_POST["nombrepa2"]."',".$dianacpa2.",".$mesnacpa2.",".$anionacpa2.",".$edadpa2.",".$edadmespa2.",'".$_POST["sexopa2"]."','".$_POST["curppa2"]."','".$_POST["civilpa2"]."','".$_POST["estudiospa2"]."','".$_POST["laboralpa2"]."','".$_POST["nacionalidadpa2"]."','".$_POST["entidadnacpa2"]."','".$_POST["documentopa2"]."','".$_POST["callepa2"]."','".$_POST["calle1pa2"]."','".$_POST["calle2pa2"]."','".$_POST["numextpa2"]."','".$_POST["numintpa2"]."','".$_POST["manzanapa2"]."','".$_POST["lotepa2"]."','".$_POST["deptopa2"]."','".$_POST["coloniapa2"]."','".$_POST["cppa2"]."','".$_POST["localidadpa2"]."','".$_POST["municipiopa2"]."','".$_POST["referenciapa2"]."','".$_POST["entidadpa2"]."','".$_POST["telcasapa2"]."','".$_POST["telcelpa2"]."','".$_POST["correo1pa2"]."','".$_POST["correo2pa2"]."')";
  
  if ($conn->query($query) === TRUE) {
      echo "Datos guardados correctamente.";
  } else {
      echo "Error: " . $query . "<br>" . $conn->error;
  }
  
  // Cerrar la conexión
  $conn->close();

  echo '<hr>
  <h1>
  Da click en las letras azules para descargar e imprimir tu formato de inscripción 
  <hr> <a href="docs_himno/'.$nombre.'"><img src=manita.png width=60 height=60>'.$nombre.'</a> 
  <br>y lleválo impreso el día de la inscripción a la escuela junto con este otro documento:
  <br><a href="docs_himno/aec_inscripcion_h.docx"><img src=manita.png width=60 height=60>ACUERDO ESCOLAR DE CONVIVENCIA (AEC)</a> 
  </h1>
  <hr>
  

  ';
  die();

}


?>
 <!DOCTYPE html>
<html>

<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
  <meta name="description" content="Start your development with a Dashboard for Bootstrap 4.">
  <meta name="author" content="Creative Tim">
  <title>Formato de inscripcion</title>
  <!-- Favicon -->
  <link rel="icon" href="assets/img/brand/favicon.png" type="image/png">
  <!-- Fonts -->
  <link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Open+Sans:300,400,600,700">
  <!-- Icons -->
  <link rel="stylesheet" href="assets/vendor/nucleo/css/nucleo.css" type="text/css">
  <link rel="stylesheet" href="assets/vendor/@fortawesome/fontawesome-free/css/all.min.css" type="text/css">
  <!-- Page plugins -->
  <link rel="stylesheet" href="assets/vendor/select2/dist/css/select2.min.css">
  <link rel="stylesheet" href="assets/vendor/quill/dist/quill.core.css">
  <!-- Argon CSS -->
  <link rel="stylesheet" href="assets/css/argon.css?v=1.1.0" type="text/css">

  <script>
  function validateForm() {
    var inputs = document.getElementsByTagName('input');
    var selects = document.getElementsByTagName('select');
    var filled = true;

    for (var i = 0; i < inputs.length; i++) {
      if (inputs[i].type !== 'radio' && inputs[i].type !== 'submit') {
        var fieldName = inputs[i].getAttribute('name');
        if (fieldName !== 'lote' && fieldName !== 'manzana' && fieldName !== 'depto' && fieldName !== 'numint' 
        && fieldName !== 'lotepa' && fieldName !== 'manzanapa' && fieldName !== 'deptopa' && fieldName !== 'numintpa' 
        && fieldName !== 'lotepa2' && fieldName !== 'manzanapa2' && fieldName !== 'deptopa2' && fieldName !== 'numintpa2'
        && fieldName !== 'gradoherm' && fieldName !== 'grupoherm'
    && fieldName !== 'apellido1herm' && fieldName !== 'apellido2herm'
    && fieldName !== 'redsocial' && fieldName !== 'tiporeded'
    && fieldName !== 'correo2' && fieldName !== 'correo2pa' && fieldName !== 'correo2pa2'
    && fieldName !== 'localidad' && fieldName !== 'entidad'
    && fieldName !== 'localidadpa' && fieldName !== 'entidadpa'
    && fieldName !== 'localidadpa2' && fieldName !== 'entidadpa2'
    && fieldName !== 'telcasa' && fieldName !== 'telcel'
    && fieldName !== 'telcasapa' && fieldName !== 'telcelpa'
    && fieldName !== 'telcasapa2' && fieldName !== 'telcelpa2'
    && fieldName !== 'nombreherm' && fieldName !== 'apellido2herm' ) {
          if (inputs[i].value === '') {
            filled = false;
            break;
          }
        }
      }
    }

    for (var j = 0; j < selects.length; j++) {
      if (selects[j].value === '') {
        filled = false;
        break;
      }
    }

    if (!filled) {
      alert('Por favor, complete todos los campos obligatorios.');
    }

    return filled;
  }
</script>

</head>

<body>


<form method="post" onsubmit="return validateForm();">

    <!-- Header -->
    <!-- Header -->
    <div class="header bg-primary pb-6">
      <div class="container-fluid">
        <div class="header-body">
          <div class="row align-items-center py-4">
            <div class="col-lg-6 col-7">
              <h6 class="h2 text-white d-inline-block mb-0">Formato de inscripción</h6>
              <nav aria-label="breadcrumb" class="d-none d-md-inline-block ml-md-4">
                <ol class="breadcrumb breadcrumb-links breadcrumb-dark">
                  <li class="breadcrumb-item"><a href="#"><i class="fas fa-home"></i></a></li>
                  <li class="breadcrumb-item"><a href="#">Escuela Secundario HIMNO NACIONAL</a></li>
                  <li class="breadcrumb-item active" aria-current="page">Inscripción</li>
                </ol>
              </nav>
            </div>
            <div class="col-lg-6 col-5 text-right">
              <input type="submit" value="Guardar">
              
            </div>
          </div>
        </div>
      </div>
    </div>
<!-- Page content -->
<div class="container-fluid mt--6">
      <div class="row">
        <div class="col-lg-12">
          <div class="card-wrapper">

            <!-- ALUMNO(A): -->
            <div class="card">
              <!-- Card header -->
              <div class="card-header">
                <h3 class="mb-0">ALUMNO(A): </h3>
              </div>
              <!-- Card body -->
              <div class="card-body">
                

                  <!-- datos del alumno -->
                  <div class="row">

                  <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="apellido1" placeholder="PRIMER APELLIDO" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="apellido2" placeholder="SEGUNDO APELLIDO" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">

                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="nombre" placeholder="NOMBRE(S)" type="text">
                        </div>
                      </div>
                    </div>


                    <div class="col-md-2">
                    <div class="form-group">
                        <!--label class="form-control-label" for="exampleDatepicker">FECHA DE NACIMIENTO</label-->
                        <input class="form-control datepicker" name="fechanac" placeholder="FECHA DE NACIMIENTO" type="text">
                      </div>
                    </div>

                    <div class="col-md-2">

                    <div class="custom-control custom-radio mb-3">
                        <input name="sexo" value="H" class="custom-control-input" id="customRadio5" type="radio">
                        <label class="custom-control-label" for="customRadio5">Hombre</label>
                      </div>
                      <div class="custom-control custom-radio mb-3">
                        <input name="sexo" value="M" class="custom-control-input" id="customRadio6" type="radio">
                        <label class="custom-control-label" for="customRadio6">Mujer</label>
                      </div>

                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="curp" placeholder="CURP" type="text">
                        </div>
                      </div>
                    </div>


                  </div>


                  <div class="row">

<div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="peso" placeholder="PESO" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="talla" placeholder="TALLA" type="text">
      </div>
    </div>
  </div>

<div class="col-md-2">
  <label class="form-control-label" for="lentes">LENTES</label>
  <div class="custom-control custom-radio mb-3">
    <input name="lentes" value="SI" class="custom-control-input" id="lentessi" type="radio">
    <label class="custom-control-label" for="lentessi">SI</label>
  </div>
  <div class="custom-control custom-radio mb-3">
    <input name="lentes" value="NO" class="custom-control-input" id="lentesno" type="radio">
    <label class="custom-control-label" for="lentesno">NO</label>
  </div>

</div>

<div class="col-md-2">
  <label class="form-control-label" for="lentes">ZAPATOS ORTOPEDICOS</label>
  <div class="custom-control custom-radio mb-3">
    <input name="zapatos" value="SI" class="custom-control-input" id="zapsi" type="radio">
    <label class="custom-control-label" for="zapsi">SI</label>
  </div>
  <div class="custom-control custom-radio mb-3">
    <input name="zapatos" value="NO" class="custom-control-input" id="zapno" type="radio">
    <label class="custom-control-label" for="zapno">NO</label>
  </div>

</div>

<div class="col-md-2">
  <label class="form-control-label" for="lentes">CARTILLA DE VACUNACION</label>
  <div class="custom-control custom-radio mb-3">
    <input name="cartilla" value="SI" class="custom-control-input" id="cartsi" type="radio">
    <label class="custom-control-label" for="cartsi">SI</label>
  </div>
  <div class="custom-control custom-radio mb-3">
    <input name="cartilla" value="NO" class="custom-control-input" id="cartno" type="radio">
    <label class="custom-control-label" for="cartno">NO</label>
  </div>

</div>

<div class="col-md-2">
  <label class="form-control-label" for="lentes">VACUNAS COMPLETAS</label>
  <div class="custom-control custom-radio mb-3">
    <input name="vacunas" value="SI" class="custom-control-input" id="vacsi" type="radio">
    <label class="custom-control-label" for="vacsi">SI</label>
  </div>
  <div class="custom-control custom-radio mb-3">
    <input name="vacunas" value="NO" class="custom-control-input" id="vacno" type="radio">
    <label class="custom-control-label" for="vacno">NO</label>
  </div>

</div>

</div>
<div class="row">

  <div class="col-md-2">
    <div class="form-group">
      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="nacionalidad" placeholder="NACIONALIDAD" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-2">
    <div class="form-group">
      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="entidadnac" placeholder="ENTIDAD DE NACIMIENTO" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-2">
    <div class="form-group">
                    <label class="form-control-label" for="serviciomed">TIPO DE SERVICIO MEDICO</label>
                    <input class="form-control" name="serviciomed" id="serviciomed" list="servicios_medicos" placeholder="TIPO DE SERVICIO MEDICO" type="text">
                    <datalist id="servicios_medicos">
                      <option value="IMSS">
                      <option value="ISSSTE">
                      <option value="ISSEMYM">
                      <option value="ISEM">
                      <option value="PARTICULAR">
                    </datalist>
    </div>

  </div>


  <div class="col-md-1">
    <div class="form-group">
                    <label class="form-control-label" for="grado">GRADO</label>
                    <select class="form-control" name="grado" id="serviciomed">
                      <option>1</option>
                      <option>2</option>
                      <option>3</option>

                    </select>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
                    <label class="form-control-label" for="grupo">GRUPO</label>
                    <select class="form-control" name="grupo" id="serviciomed">
                    <option>A</option>
                    <option>B</option>
                    <option>C</option>
                    <option>D</option>
                    <option>E</option>
                    <option>F</option>
                    <option>G</option>
                    <option>H</option>
                    <option>I</option>  
                    <option>J</option>
                    <option>K</option>
                    <option>L</option>


                    </select>
    </div>
  </div>


</div>


                
              </div>
            </div>



            <!-- DOMICILIO DEL ALUMNO(A): -->
            <div class="card">
              <!-- Card header -->
              <div class="card-header">
                <h3 class="mb-0">DOMICILIO DEL ALUMNO(A): </h3>
              </div>
              <!-- Card body -->
              <div class="card-body">
                

                  <!-- Input groups with icon -->
                  <div class="row">

                  <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="calle" placeholder="CALLE" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="entrecalle1" placeholder="ENTRE LA CALLE" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">

                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="entrecalle2" placeholder="Y LA CALLE" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="colonia" placeholder="COLONIA" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="municipio" placeholder="MUNICIPIO" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="localidad" placeholder="LOCALIDAD" type="text">
                        </div>
                      </div>
                    </div>


                  </div><!-- fin de row-->


                  <div class="row">

<div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="numext" placeholder="#Ext." type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="numint" placeholder="#Int." type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="manzana" placeholder="Manzana" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="lote" placeholder="Lote" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="depto" placeholder="Depto" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <!--input class="form-control" name="talla" placeholder="TALLA" type="text"-->
      </div>
    </div>
  </div>

<div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="cp" placeholder="CODIGO POSTAL" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-4">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="referencia" placeholder="OTRA REFERENCIA  (Escuela, Iglesia, Tienda etc.)" type="text">
                        </div>
                      </div>
                    </div>

</div>
<div class="row">

  <div class="col-md-2">
    <div class="form-group">
      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="entidad" placeholder="ENTIDAD" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-2">
    <div class="form-group">
      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="telcasa" placeholder="TEL. CASA" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-2">
    <div class="form-group">
      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="telcel" placeholder="TEL. CELULAR" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-4">
    <div class="form-group">
      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="redsocial" placeholder="RED SOCIAL PRINCIPAL: NOMBRE O ENLACE" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-2">
    <div class="form-group">
      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="tipored" placeholder="TIPO (FACEBOOK, INSTAGRAM, ETC.) " type="text">
      </div>
    </div>
  </div>


  </div><!-- fin de row-->

                  <!-- inicia row correos -->
                  <div class="row">

                  <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="correo1" placeholder="EMAIL PERSONAL " type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="correo2" placeholder="EMAIL NUEVAESCUELA.MX" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
  <label class="form-control-label" for="lentes">¿TIENE HERMANOS INSCRITOS ACTUALMENTE EN ESTA ESCUELA?</label>
  <div class="custom-control custom-radio mb-3">
    <input name="hermanos" value="SI" class="custom-control-input" id="hermsi" type="radio">
    <label class="custom-control-label" for="hermsi">SI</label>
  </div>
  <div class="custom-control custom-radio mb-3">
    <input name="hermanos" value="NO" class="custom-control-input" id="hermno" type="radio">
    <label class="custom-control-label" for="hermno">NO</label>
  </div>

</div>

<div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="gradoherm" placeholder="GRADO" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="grupoherm" placeholder="GRUPO" type="text">
      </div>
    </div>
  </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="apellido1herm" placeholder="PRIMER APELLIDO" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="apellido2herm" placeholder="SEGUNDO APELLIDO" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="nombreherm" placeholder="NOMBRE(S)" type="text">
                        </div>
                      </div>
                    </div>

                  </div><!-- fin de row-->
             
              </div>
            </div>

            <!-- DOMICILIO DEL ALUMNO(A): -->
            <div class="card">
              <!-- Card header -->
              <div class="card-header">
                <h3 class="mb-0">DATOS DEL PADRE, MADRE DE FAMILIA O TUTOR: </h3>
              </div>
              <!-- Card body -->
              <div class="card-body">
                

                  <!-- PADRE O TUTOR -->
                  <!-- PA -->
                  <div class="row">

                  <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="parentesco" placeholder="PARENTESCO" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="apellido1pa" placeholder="PRIMER APELLIDO" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="apellido2pa" placeholder="SEGUNDO APELLIDO" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">

                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="nombrepa" placeholder="NOMBRE(S)" type="text">
                        </div>
                      </div>
                    </div>


                    <div class="col-md-2">
                    <div class="form-group">
                        <!--label class="form-control-label" for="exampleDatepicker">FECHA DE NACIMIENTO</label-->
                        <input class="form-control datepicker" name="fechanacpa" placeholder="FECHA DE NACIMIENTO" type="text">
                      </div>
                    </div>

                    <div class="col-md-1">

                    <div class="custom-control custom-radio mb-3">
                        <input name="sexopa" value="H" class="custom-control-input" id="customRadio7" type="radio">
                        <label class="custom-control-label" for="customRadio7">Hombre</label>
                      </div>
                      <div class="custom-control custom-radio mb-3">
                        <input name="sexopa" value="M" class="custom-control-input" id="customRadio8" type="radio">
                        <label class="custom-control-label" for="customRadio8">Mujer</label>
                      </div>

                    </div>

                    <div class="col-md-1">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="curppa" placeholder="CURP" type="text">
                        </div>
                      </div>
                    </div>


                  </div><!-- fin de row-->

                  <div class="row"><!-- civil -->

                  <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="civil" placeholder=" ESTADO CIVIL" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="estudios" placeholder="GRADO MÁXIMO DE ESTUDIOS" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">

                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="laboral" placeholder="SITUACIÓN LABORAL (con / sin empleo) " type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="nacionalidadpa" placeholder="NACIONALIDAD" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="entidadnacpa" placeholder="ENTIDAD DE NACIMIENTO" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="documento" placeholder="TIPO DE DOCUMENTO OFICIAL" type="text">
                        </div>
                      </div>
                    </div>


                  </div><!-- fin de row civil-->

                  <div class="row"><!--calle-->

                  <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="callepa" placeholder="CALLE" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="calle1pa" placeholder="ENTRE LA CALLE" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">

                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="calle2pa" placeholder="Y LA CALLE" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="coloniapa" placeholder="COLONIA" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="municipiopa" placeholder="MUNICIPIO" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="localidadpa" placeholder="LOCALIDAD" type="text">
                        </div>
                      </div>
                    </div>



                  </div><!-- fin de row -->
                  <div class="row">

<div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="numextpa" placeholder="#Ext." type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="numintpa" placeholder="#Int." type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="manzanapa" placeholder="Manzana" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="lotepa" placeholder="Lote" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="deptopa" placeholder="Depto" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <!--input class="form-control" name="talla" placeholder="TALLA" type="text"-->
      </div>
    </div>
  </div>

<div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="cppa" placeholder="CODIGO POSTAL" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-4">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="referenciapa" placeholder="OTRA REFERENCIA  (Escuela, Iglesia, Tienda etc.)" type="text">
                        </div>
                      </div>
                    </div>

</div><!-- fin de row -->

<div class="row"><!--entidad y correos-->

<div class="col-md-2">
    <div class="form-group">
      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="entidadpa" placeholder="ENTIDAD" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="telcasapa" placeholder="TEL CASA" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="telcelpa" placeholder="TEL CELULAR" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-2">
    <div class="form-group">

      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="correo1pa" placeholder="EMAIL PRINCIPAL" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-2">
    <div class="form-group">
      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="correo2pa" placeholder="EMAIL ALTERNO" type="text">
      </div>
    </div>
  </div>

<div class="col-md-2">
  <label class="form-control-label" for="lentes">¿EL ALUMNO TIENE NECESIDADES EDUCATIVAS ESPECIALES?</label>
  <div class="custom-control custom-radio mb-3">
    <input name="especial" value="SI" class="custom-control-input" id="espesi" type="radio">
    <label class="custom-control-label" for="espesi">SI</label>
  </div>
  <div class="custom-control custom-radio mb-3">
    <input name="especial" value="NO" class="custom-control-input" id="espeno" type="radio">
    <label class="custom-control-label" for="espeno">NO</label>
  </div>
</div>

<div class="col-md-2">
  <label class="form-control-label" for="lentes">¿EL ALUMNO REQUIERE HERRAMIENTAS DE APOYO PARA EL APRENDIZAJE?</label>
  <div class="custom-control custom-radio mb-3">
    <input name="apoyo" value="SI" class="custom-control-input" id="apoyosi" type="radio">
    <label class="custom-control-label" for="apoyosi">SI</label>
  </div>
  <div class="custom-control custom-radio mb-3">
    <input name="apoyo" value="NO" class="custom-control-input" id="apoyono" type="radio">
    <label class="custom-control-label" for="apoyono">NO</label>
  </div>
</div>

<div class="col-md-2">
  <label class="form-control-label" for="lentes">¿PERTENECE A ALGÚN GRUPO INDIGENA?	</label>
  <div class="custom-control custom-radio mb-3">
    <input name="indigena" value="SI" class="custom-control-input" id="indisi" type="radio">
    <label class="custom-control-label" for="indisi">SI</label>
  </div>
  <div class="custom-control custom-radio mb-3">
    <input name="indigena" value="NO" class="custom-control-input" id="indino" type="radio">
    <label class="custom-control-label" for="indino">NO</label>
  </div>
</div>



</div><!-- fin de row calle-->

</div><!-- fin card body -->
</div><!-- fin card header -->

            <!-- DATOS DE LA SEGUNDA PERSONA RESPONSABLE DEL ALUMNO (A): -->
            <div class="card">
              <!-- Card header -->
              <div class="card-header">
                <h3 class="mb-0">DATOS DE LA SEGUNDA PERSONA RESPONSABLE DEL ALUMNO (A): </h3>
              </div>
              <!-- Card body -->
              <div class="card-body">
                

                  <!-- SEGUNDO TUTOR -->
                  <!-- PA2 -->
                  <div class="row">

                  <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="parentescopa2" placeholder="PARENTESCO" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="apellido1pa2" placeholder="PRIMER APELLIDO" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="apellido2pa2" placeholder="SEGUNDO APELLIDO" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">

                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="nombrepa2" placeholder="NOMBRE(S)" type="text">
                        </div>
                      </div>
                    </div>


                    <div class="col-md-2">
                    <div class="form-group">
                        <!--label class="form-control-label" for="exampleDatepicker">FECHA DE NACIMIENTO</label-->
                        <input class="form-control datepicker" name="fechanacpa2" placeholder="FECHA DE NACIMIENTO" type="text">
                      </div>
                    </div>

                    <div class="col-md-1">

                    <div class="custom-control custom-radio mb-3">
                        <input name="sexopa2" value="H" class="custom-control-input" id="customRadio9" type="radio">
                        <label class="custom-control-label" for="customRadio9">Hombre</label>
                      </div>
                      <div class="custom-control custom-radio mb-3">
                        <input name="sexopa2" value="M" class="custom-control-input" id="customRadio10" type="radio">
                        <label class="custom-control-label" for="customRadio10">Mujer</label>
                      </div>

                    </div>

                    <div class="col-md-1">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="curppa2" placeholder="CURP" type="text">
                        </div>
                      </div>
                    </div>


                  </div><!-- fin de row-->

                  <div class="row"><!-- civil -->

                  <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="civilpa2" placeholder=" ESTADO CIVIL" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="estudiospa2" placeholder="GRADO MÁXIMO DE ESTUDIOS" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">

                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="laboralpa2" placeholder="SITUACIÓN LABORAL (con / sin empleo) " type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="nacionalidadpa2" placeholder="NACIONALIDAD" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="entidadnacpa2" placeholder="ENTIDAD DE NACIMIENTO" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="documentopa2" placeholder="TIPO DE DOCUMENTO OFICIAL" type="text">
                        </div>
                      </div>
                    </div>


                  </div><!-- fin de row civil-->

                  <div class="row"><!--calle-->

                  <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="callepa2" placeholder="CALLE" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="calle1pa2" placeholder="ENTRE LA CALLE" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">

                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="calle2pa2" placeholder="Y LA CALLE" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="coloniapa2" placeholder="COLONIA" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="municipiopa2" placeholder="MUNICIPIO" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="localidadpa2" placeholder="LOCALIDAD" type="text">
                        </div>
                      </div>
                    </div>



                  </div><!-- fin de row -->
                  <div class="row">

<div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="numextpa2" placeholder="#Ext." type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="numintpa2" placeholder="#Int." type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="manzanapa2" placeholder="Manzana" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="lotepa2" placeholder="Lote" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <input class="form-control" name="deptopa2" placeholder="Depto" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">

        <!--input class="form-control" name="talla" placeholder="TALLA" type="text"-->
      </div>
    </div>
  </div>

<div class="col-md-2">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="cppa2" placeholder="CODIGO POSTAL" type="text">
                        </div>
                      </div>
                    </div>

                    <div class="col-md-4">
                      <div class="form-group">
                        <div class="input-group input-group-merge">
                          <div class="input-group-prepend">
                            <span class="input-group-text"><i class="fas fa-user"></i></span>
                          </div>
                          <input class="form-control" name="referenciapa2" placeholder="OTRA REFERENCIA  (Escuela, Iglesia, Tienda etc.)" type="text">
                        </div>
                      </div>
                    </div>

</div><!-- fin de row -->

<div class="row"><!--entidad y correos-->

<div class="col-md-2">
    <div class="form-group">
      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="entidadpa2" placeholder="ENTIDAD" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="telcasapa2" placeholder="TEL CASA" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-1">
    <div class="form-group">
      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="telcelpa2" placeholder="TEL CELULAR" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-2">
    <div class="form-group">

      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="correo1pa2" placeholder="EMAIL PRINCIPAL" type="text">
      </div>
    </div>
  </div>

  <div class="col-md-2">
    <div class="form-group">
      <div class="input-group input-group-merge">
        <div class="input-group-prepend">
          <span class="input-group-text"><i class="fas fa-user"></i></span>
        </div>
        <input class="form-control" name="correo2pa2" placeholder="EMAIL ALTERNO" type="text">
      </div>
    </div>
  </div>

</div><!-- fin de row calle-->

</div><!-- fin card body -->
</div><!-- fin card header -->

</div>          
          </div>
        </div>
        
        <div class="row">
          <div class="col-lg-12">
            <div class="card-wrapper">
              <div class="col-lg-6 col-5 text-right">
                <input type="submit" value="Guardar">
              </div>
            </div>
          </div>
        </div>
      </div>
      <!-- Footer -->
      <footer class="footer pt-0">
        <div class="row align-items-center justify-content-lg-between">
          <div class="col-lg-6">
            <div class="copyright text-center text-lg-left text-muted">
              &copy; 2025 Escuela Secundaria HIMNO NACIONAL
            </div>
          </div>

        </div>
        </footer>
    </div>
  </div> 
  <!-- Argon Scripts -->
  <!-- Core -->
  <script src="assets/vendor/jquery/dist/jquery.min.js"></script>
  <script src="assets/vendor/bootstrap/dist/js/bootstrap.bundle.min.js"></script>
  <script src="assets/vendor/js-cookie/js.cookie.js"></script>
  <script src="assets/vendor/jquery.scrollbar/jquery.scrollbar.min.js"></script>
  <script src="assets/vendor/jquery-scroll-lock/dist/jquery-scrollLock.min.js"></script>
  <!-- Optional JS -->
  <script src="assets/vendor/select2/dist/js/select2.min.js"></script>
  <script src="assets/vendor/bootstrap-datepicker/dist/js/bootstrap-datepicker.min.js"></script>
  <script src="assets/vendor/nouislider/distribute/nouislider.min.js"></script>
  <script src="assets/vendor/quill/dist/quill.min.js"></script>
  <script src="assets/vendor/dropzone/dist/min/dropzone.min.js"></script>
  <script src="assets/vendor/bootstrap-tagsinput/dist/bootstrap-tagsinput.min.js"></script>
  <!-- Argon JS -->
  <script src="assets/js/argon.js?v=1.1.0"></script>
  <!-- Demo JS - remove this in your project -->
  <script src="assets/js/demo.min.js"></script>

  </form>  
</body>

</html>