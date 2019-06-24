
/*Use code below to export many xlsx-files in one step*/

/*Macros for deleting single file - copy this without a change
	>>> If you already had files with the same name at the directory, 
		you'd better delete them before export a newest version */
%macro delete_file(directory, name);
	filename x_name "&directory\&name..xlsx";
  	data _null_ ; 
    rc = fdelete('x_name') ; 
  	run ; 
%mend;

/*Macros for exporting single file - copy this without a change */
%macro export_xlsx(directory, name, source_table);
proc export data = &source_table
dbms=excel 
outfile = "&directory\&name..xlsx"
label
replace;    
run;
%mend;

/* Macros for exporting many files in one step - copy this and make a change by comments below 
	>>> macro step uses two macro steps upper ^^^ */
	/* All data for export in files should be collected in one table "common_table" .
	   Files will be sorted by "filter" column */
%macro export_files(filter,common_table,directory);
%let num = 0;    

	data common_table; 
	set &common_table.;
	run;

	proc sql noprint;
	select count(*) = 0 into :num 
	from common_table;
	quit;

%do %while (&num = 0); 

	proc sort data = common_table; 
	by &filter.;
	run;

	data single_file (drop = filter_column)
    common_table (drop = filter_column);
	set common_table;
	retain filter_column;

		if _n_ = 1 then do; 
		 	filter_column = compress(&filter.); 
 			output single_file; 
		end;

		else if filter_column = compress(&filter.) then do;
  			output single_file;
  			filter_column = compress(&filter.);
		end;

  		else output common_table;
	run;

	proc sql noprint;
	select compress(&filter.) into:by 
	from single_file;
	quit;

%if &num = 0 %then %do;              
%let str = %sysfunc(compress(&by.));

%delete_file(&directory., &str._new);

%export_xlsx(&directory., &str._new, single_file);

%end;

	proc sql noprint;
	select count(*) = 0 into: num from common_table;
	quit;

%end;
	proc sql noprint; 
	drop table common_table, common_table; 
	quit;

%mend;
/*Example of macros invocation*/
%export_files(Team, sashelp.baseball, C:\Users\Noname\example\);
