CREATE TABLE `projeto-id.dataset_name.table_name`
 (
   ano INT64,
   mes INT64,
   estado string,
   produto string,
   unidade string,
   vol_demanda_m3 float64,
   timestamp_captura timestamp
 )
 PARTITION BY DATE(timestamp_captura)