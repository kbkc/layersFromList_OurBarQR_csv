//layerIniName -> Put here first ASCII code (A=65, a=97) 
//numberOfLayers -> Put here how many layer do you need 
myfunc();


function myfunc()
{

var a;
DrawText2Layers();

// drawBarCode(EAN2Bin("5055379414694"));
}



function DrawText2Layers()
{
//Apply to myDoc the active document 
var myDoc = app.activeDocument; 
//define firts caracter and how many layers do you need 
var arr = [
"941446",
"941447",
"941448",
"941449",
"941450",
"941451",
"941452",
"941453",
"941454",
"941455",
"941456",
"941457",
"941458",
"941459",
"941460",
"941461",
"941462",
"941463",
"941464",
"941465",
"941466",
"941467",
"941468",
"941469"
];

arr2 =[
"Bitter Coffee",
"Flamenco",
"Frozen Mango",
"Cosmopolitan ",
"Double Whiskey",
"Sunny Rays",
"Deepness ",
"Sangria",
"First Love",
"Hawaiian Pink",
"Cerise Hood",
"Merlot",
"Scarlet Sails",
"Mandarin Red",
"Damask Rose",
"Heritage Terracotta",
"Passionate Apricot",
"Magenta Siren",
"Beaujolais ",
"Burgundy Dress",
"Burning Flames",
"The Abyss",
"Iceberg",
"Glamorous Peach"
];

arr3 = [
"5055379414465",
"5055379414472",
"5055379414489",
"5055379414496",
"5055379414502",
"5055379414519",
"5055379414526",
"5055379414533",
"5055379414540",
"5055379414557",
"5055379414564",
"5055379414571",
"5055379414588",
"5055379414595",
"5055379414601",
"5055379414618",
"5055379414625",
"5055379414632",
"5055379414649",
"5055379414656",
"5055379414663",
"5055379414670",
"5055379414687",
"5055379414694"
];

//Create the layers 
for(var i=0; i<arr.length; i++) 
	{ 
		var myLayer = myDoc.layers.add(); 
		myLayer.name = arr[i];

		var art=myDoc.textFrames.add();
		//art.Position = Array(200, 200);
		art.top = 0;
		art.left=400;
		art.contents = arr[i];

		var tname=myDoc.textFrames.add();
		//tname.Position = Array(300, 300);
		tname.top = 0;
		tname.left=450;
		tname.contents = arr2[i];


        drawBarCode(EAN2Bin(arr3[i]));

	}
} 








function drawBarCode(bc)
{
	var i, j,k,col,w,h;
	var barColor = new CMYKColor();
	barColor.black = 100;
	barColor.cyan = 0;
	barColor.magenta = 0;
	barColor.yellow = 0;

	var bkgColor = new CMYKColor();
	bkgColor.black = 0;
	bkgColor.cyan = 0;
	bkgColor.magenta = 0;
	bkgColor.yellow = 0;
           /*
          h:=StrToInt(Form1.eHeight.Text); // pic. height  ������� �������
          Form1.Image1.Width := length(bc);
          Form1.Image1.Height:= h;
		  */
		  var mm=1//2.834645669;
		  var x1,y1,w1,h1;
		  x1=100*mm;
		  y1=1*mm;
		  w1=bc.length*mm;
		  h1=10*mm;
		  
		  
		 
		 var docRef = app.activeDocument;  
		  
		  var hh=3;
		  
		  var rectbkg = docRef.pathItems.rectangle( x1+hh, y1-hh, (w1+hh*2), (h1*mm +hh*2));
		  rectbkg.stroked = false;
	      rectbkg.filled = true;
          rectbkg.fillColor = bkgColor;
		  
		  
		  
          var EANGroup = docRef.groupItems.add(); 

          for(var i = 0;i< bc.length;i++)
        	{
			 if (bc[i]=='1')
			 {
		  	  j=i;k=0;
			  while(bc[i]=='1'){
			    i++;
				k++;
			   }
			  	 var rect = EANGroup.pathItems.rectangle( x1, y1+j*mm, k*mm, h1 ); // ��������������� � 1 ��������
			     //var rect = docRef.pathItems.rectangle( 100*mm, j*mm, k*mm, 10*mm ); //�� ��������������� ��������������
	             rect.stroked = false;
	             rect.filled = true;
                 rect.fillColor = barColor;
				 
			  }
         }
}







function EAN2Bin(strEANCode) 
{
      
var K;
    var strAux, strExit, strCode,e;
  //  strEANCode="5055379414465";
   // strEANCode=strEANCode.trim();
    strAux = strEANCode.substring(1, strEANCode.length);
    switch(strEANCode[0])
    {
        case '0' : strCode = '000000';break;
        case '1' : strCode = '001011';break;
        case '2' : strCode = '001101';break;
        case '3' : strCode = '001110';break;
        case '4' : strCode = '010011';break;
        case '5' : strCode = '011001';break;
        case '6' : strCode = '011100';break;
        case '7' : strCode = '010101';break;
        case '8' : strCode = '010110';break;
        case '9' : strCode = '011010';break;
    }
    strExit = '101';
    // ������ � ������ �������� ����.
    for( K = 0; K<strAux.length/2;K++){
        switch( strAux[K] ){
            case '0': e= '0100111'; if(strCode[K] == '0') { e= '0001101'  ;}
                strExit = strExit + e;  break;
            case '1': e= '0110011'; if(strCode[K] == '0') { e= '0011001';}
                strExit = strExit + e;  break;
            case '2':e='0011011'; if(strCode[K] == '0') { e= '0010011';}
                strExit = strExit + e;  break;
            case '3':e='0100001'; if(strCode[K] == '0') { e='0111101';}
                strExit = strExit + e;  break;
            case '4':e= '0011101'; if(strCode[K] == '0') { e= '0100011';}
                strExit = strExit + e;  break;
            case '5':e= '0111001'; if(strCode[K] == '0') { e= '0110001';}
                strExit = strExit + e;  break;
            case '6':e= '0000101'; if(strCode[K] == '0'){  e= '0101111';}
                strExit = strExit + e;  break;
            case '7':e= '0010001'; if(strCode[K] == '0') { e= '0111011';}
                strExit = strExit + e;  break;
            case '8':e= '0001001'; if(strCode[K] == '0')  {e= '0110111';}
                strExit = strExit + e;  break;
            case '9':e= '0010111'; if(strCode[K] == '0')  {e=  '0001011'  ;}
                strExit = strExit + e;  break;
        }
    }
//alert(strAux.length / 2 + 1);
    strExit =strExit + '01010';
	//strExit =strExit + '00000';
    for (K = (strAux.length / 2 ); K<=( strAux.length);K++)
    {
    switch (strAux[K]){
        case '0': strExit =  strExit + '1110010';break;
        case '1': strExit =  strExit + '1100110';break;
        case '2': strExit =  strExit + '1101100';break;
        case '3': strExit =  strExit + '1000010';break;
        case '4': strExit =  strExit + '1011100';break;
        case '5': strExit =  strExit + '1001110';break;
        case '6': strExit =  strExit + '1010000';break;
        case '7': strExit =  strExit + '1000100';break;
        case '8': strExit =  strExit + '1001000';break;
        case '9': strExit =  strExit + '1110100';break;
    }
    }
    strExit =strExit+ '101';
	  return  strExit;
}


