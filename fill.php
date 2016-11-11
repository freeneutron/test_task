<?php
// phpinfo();

ini_set('max_execution_time', 300);

new Fill_table;

class Fill_table{
  private $xlsx_path_default= 'xml/gate-1041__in.xlsx';
  private $sqlite_path= 'data/base.sqlite';
  private $table_name= 'gate';
  private $set_name_list= array('manager','progect');
  private $index_name= 'idp';
  private $col_name= array(
    'idp'=>'B',
    'manager'=>'C',
    'progect'=>'D',
  );

  public function __construct(){
    $xlsx_path= isset($argv[1])&& file_exists($argv[1])? $argv[1]: $this->xlsx_path_default;
    $this->xlsx= new Xlsx_file($xlsx_path);
    $col_b= $this->xlsx->get_col($this->col_name[$this->index_name]);
    $col_b_index= $this->get_index($col_b);
    $test_init_sql_request= $this->create_test_init_sql_request($col_b_index,array_keys($this->col_name));
    $data= $this->get_data($col_b_index,$this->index_name);
    $this->add_data($data,$col_b_index);
    $this->xlsx->save($this->create_target_path($xlsx_path));
    echo"ok";
  }
  private function create_target_path($path){
    $pathinfo= pathinfo($path);
    $name= $pathinfo['filename'];
    $name.= ".target";
    $target_path= "{$pathinfo['dirname']}/{$name}.{$pathinfo['extension']}";
    return $target_path;
  }
  private function get_index($array){
    $index= array();
    foreach($array as $name=>$value){
      if(!isset($index[$value])){
        $index[$value]= array();
      }
      $index[$value][]= $name;
    }
    return $index;
  }
  private function add_data($data,$col_b_index){
    $shared_string_list= $this->get_shared_string_list($data,$this->set_name_list);
    $shared_string_index= $this->xlsx->add_shared_string($shared_string_list);
    foreach($col_b_index as $key=>$row_list){
      if(isset($data[$key])){
        foreach($row_list as $row){
          foreach($this->set_name_list as $set_name){
            $this->xlsx->set_cell($this->col_name[$set_name],$row,$shared_string_index[$data[$key][$set_name]]);
            // break(3);
          }
        }
      }
    }
  }
  private function get_shared_string_list($data,$name_list){
    $shared_string_list= array();
    foreach($data as $data_item){
      foreach($name_list as $name){
        $shared_string_list[]= $data_item[$name];
      }
    }
    return $shared_string_list;
  }
  private function get_data($index,$index_name){
    $index_list= array_keys($index);
    $sql= "SELECT * FROM {$this->table_name} WHERE $index_name in (".join(",",$index_list).");";
    $db= new SQLite3($this->sqlite_path);
    $res= $db->query($sql);
    $data= array();
    while($res_item= $res->fetchArray()){
      $data[$res_item[$index_name]]= $res_item;
    }
    return $data;
  }
  private function create_test_init_sql_request($index,$name_list){
    $sql_insert= array();
    // $index_key= array_keys($index);
    foreach($index as $key=>$item){
      $s= "$key";
      for($i=1; $i<count($name_list); $i++){
        $s.= ",'".$this->random_string(3)."'";
      }
      $sql_insert[]= "($s)";
    }
    $sql_create= array("{$name_list[0]} INTEGER PRIMARY KEY");
    for($i=1; $i<count($name_list); $i++){
      $sql_create[]= "{$name_list[$i]} TEXT";
    }
    $sql_create= "CREATE TABLE {$this->table_name} (".join(",",$sql_create).");";
    $sql_insert= "INSERT INTO {$this->table_name} (".join(",",$name_list).") values ". join(",",$sql_insert);
    return "$sql_create\r\n$sql_insert";
  }
  private function random_string($n=null){
    $al= 'abcdefghijklmnoprrstuvwxyz';
    $n= $n? $n: 10;
    $s= '';
    for($i=0; $i<$n; $i++){
      $s.= $al[rand(0,strlen($al)-1)];
    }
    return $s;
  }

}

class Xlsx_file{
  private $xml_local_path= 'xl/worksheets/sheet1.xml';
  private $xml_ss_local_path= 'xl/sharedStrings.xml';
  public function __construct($path){
    $this->load($path);
  }
  public function get_table(){
    return $this->xml;
  }
  public function set_cell($col,$row,$value){
    // $value= '3';
    $this->value= $value;
    $pattern= "/<c r=\"{$col}{$row}\"([^<]+<v>([^<]*)|[^>]+(\/>))/";

    // $col= "D";
    // $col= "E";
    // preg_match($pattern,$this->xml,$res);
    // return $res;

    $this->xml= preg_replace_callback($pattern,function($matches){
      if(count($matches)== 3){
        $res= preg_replace("/{$matches[2]}$/",$this->value,$matches[0]);
      }elseif(count($matches)== 4){
        $res= preg_replace("|{$matches[3]}$|"," t=\"s\"><v>{$this->value}</v></c>",$matches[0]);
      }else{
        $res= $matches[0];
      }
      // out(array(
      //   // '$matches'=>$matches,
      //   '$res'=>$res,
      // ));
      // $res= $matches[0];
      return $res;
    },$this->xml,1);
    return array($col,$row,$value);
  }
  public function get_col($col_name){
    $col= array();
    preg_match_all("/r=\"$col_name([0-9]+)\"[^<]+<v>([^<]*)<\/v>/",$this->xml,$res,2);
    $res= array_slice($res,1);
    foreach($res as $item){
      // $res[]= $item;
      $col[$item[1]]= $item[2];
    }
    return $col;
  }
  public function load($path){
    $this->path= $path;
    $zip= new ZipArchive;
    $res= $zip->open($path);
    if($res=== TRUE){
      $this->xml= $zip->getFromName($this->xml_local_path);
      $this->xml_ss= $zip->getFromName($this->xml_ss_local_path);
    }else{
      echo'failed, code:'.$res;
    }
    $zip->close();
  }
  public function save($path){
    $this->copy($this->path,$path);
    $zip= new ZipArchive;
    $res= $zip->open($path);
    if($res=== TRUE){
      $zip->addFromString($this->xml_local_path,$this->xml);
      $zip->addFromString($this->xml_ss_local_path,$this->xml_ss);
    }else{
      echo'failed, code:'.$res;
    }
    $zip->close();
  }
  public function add_shared_string($data){
    $index= array();
    $n= substr_count($this->xml_ss,'<si><t>');
    $s= '';
    foreach($data as $data_item){
      $index[$data_item]= $n++;
      $s.= "<si><t>$data_item</t></si>";
    }
    $this->xml_ss= str_replace("</t></si></sst>","</t></si>$s</sst>",$this->xml_ss);
    return $index;
  }
  private function copy($path_from,$path_to){
    $file= file_get_contents($path_from);
    file_put_contents($path_to,$file);
  }
}

function out($a){
  echo'<pre>';
  print_r($a);
  echo'</pre>';
}
