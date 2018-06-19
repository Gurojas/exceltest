import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver/FileSaver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'app';

  constructor(){

  }

  downloadExcel(){
    
    // se crea la hoja de calculo
    var wb = XLSX.utils.book_new();
    // se añaden propiedades como el nombre, el autor,fecha, etc
    wb.Props = {
      Title : "Excel Angular", // titulo
      Subject : "Test", // materia
      Author : "Gustafox", // autor
      CreatedDate : new Date(2017,12,19) // fecha
    };
    // se coloca nombre a la hoja
    wb.SheetNames.push("Hoja 1");
    // se añaden datos de la forma Array<Array()>()
    var ws_data = [['hello' ,'world'],
                   ['hola','mundo']
    ];
    // se agrega los datos a la hoja                 
    var ws = XLSX.utils.aoa_to_sheet(ws_data);

    //ws["!merges"] = [{ s: { c: 0, r: 0 }, e: { c: 1, r: 0 } }];
    //console.log(ws['A1']);

    // los datos ahora se asginan a la hoja
    wb.Sheets["Hoja 1"] = ws;
    // se crea la variable de salida
    var wbout = XLSX.write(wb,{bookType:'xlsx',type:'binary'});
    
    // funcion que permite convertir de binario a octeto
    function s2ab(s) { 
      var buf = new ArrayBuffer(s.length); //convert s to arrayBuffer
      var view = new Uint8Array(buf);  //create uint8array as viewer
      for (var i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF; //convert to octet
      return buf;    
    }
    
    // funcion que guarda y crea el archivo 
    saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), 'test.xlsx');

  }
}
