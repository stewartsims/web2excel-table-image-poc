//var toDataUrl = function (src, callback, outputFormat) {
//    var deferred = $.Deferred();
//    var img = new Image();
//    img.crossOrigin = 'Anonymous';
//    img.onload = function () {
//        var canvas = document.createElement('CANVAS');
//        var ctx = canvas.getContext('2d');
//        var dataURL;
//        canvas.height = this.height;
//        canvas.width = this.width;
//        ctx.drawImage(this, 0, 0);
//        deferred.resolve(dataURL = canvas.toDataURL(outputFormat));
//        callback(dataURL);
//    };
//    img.src = src;
//    if (img.complete || img.complete === undefined) {
//        img.src = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==";
//        img.src = src;
//    }
//    return deferred.promise();
//}

//function toDataUrl(url, callback) {
//    var xhr = new XMLHttpRequest();
//    xhr.responseType = 'blob';
//    xhr.onload = function() {
//        var reader = new FileReader();
//        reader.onloadend = function() {
//            callback(reader.result);
//        }
//        reader.readAsDataURL(xhr.response);
//    };
//    xhr.open('GET', url);
//    xhr.send();
//}

function toDataUrl(img, outputFormat) {
    var canvas = document.createElement('CANVAS');
    var ctx = canvas.getContext('2d');
    var dataURL;
    canvas.height = img.height;
    canvas.width = img.width;
    ctx.drawImage(img, 0, 0);
    return canvas.toDataURL(outputFormat);
}

function generateArray(table) {
    var out = [];
    var rows = table.querySelectorAll('tr');
    var ranges = [];
    var images = [];
    for (var R = 0; R < rows.length; ++R) {
        var outRow = [];
        var row = rows[R];
        var columns = row.querySelectorAll('td');
        for (var C = 0; C < columns.length; ++C) {
            var cell = columns[C];
            var colspan = cell.getAttribute('colspan');
            var rowspan = cell.getAttribute('rowspan');
            var cellValue = cell.innerText;
            if (cellValue !== "" && cellValue == +cellValue) cellValue = +cellValue;

            //Skip ranges
            ranges.forEach(function (range) {
                if (R >= range.s.r && R <= range.e.r && outRow.length >= range.s.c && outRow.length <= range.e.c) {
                    for (var i = 0; i <= range.e.c - range.s.c; ++i) outRow.push(null);
                }
            });

            //Handle Row Span
            if (rowspan || colspan) {
                rowspan = rowspan || 1;
                colspan = colspan || 1;
                ranges.push({s: {r: R, c: outRow.length}, e: {r: R + rowspan - 1, c: outRow.length + colspan - 1}});
            }
            ;


            if (cell.children.length > 0) {
                var src = cell.children[0].getAttribute("src");
                if (src !== null) {
//                    alert(toDataUrl(cell.children[0]));
                    images.push({
                        name: src,
                        data: toDataUrl(cell.children[0]).split(',')[1],
                        opts: { base64: true },
                        position: {
                            type: 'twoCellAnchor',
                            attrs: { editAs: 'oneCell' },
                            from: { col: C, row: R },
                            to: { col: C + 1, row: R + 10 }
                        }
                    });

//                    toDataUrl(src, function (base64Img) {
//                        console.log(base64Img);
//                        images.push({
//                            name: src,
//                            data: base64Img,
//                            opts: { base64: true },
//                            position: {
//                                type: 'twoCellAnchor',
//                                attrs: { editAs: 'oneCell' },
//                                from: { col: C, row: R },
//                                to: { col: C+1, row: R+1 }
//                            }
//                        });
//                    })
//                      .then(function (base64Img) {
//                        images.push({
//                            name: src,
//                            data: base64Img,
//                            opts: { base64: true },
//                            position: {
//                                type: 'twoCellAnchor',
//                                attrs: { editAs: 'oneCell' },
//                                from: { col: C, row: R },
//                                to: { col: C+1, row: R+1 }
//                            }
//                        });
//                    });
                }
            }

            //Handle Value
            outRow.push(cellValue !== "" ? cellValue : null);

            //Handle Colspan
            if (colspan) for (var k = 0; k < colspan - 1; ++k) outRow.push(null);
        }
        out.push(outRow);
    }
    return [out, ranges, images];
};

function datenum(v, date1904) {
	if(date1904) v+=1462;
	var epoch = Date.parse(v);
	return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}
 
function sheet_from_array_of_arrays(data, opts) {
	var ws = {};
	var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
	for(var R = 0; R != data.length; ++R) {
		for(var C = 0; C != data[R].length; ++C) {
			if(range.s.r > R) range.s.r = R;
			if(range.s.c > C) range.s.c = C;
			if(range.e.r < R) range.e.r = R;
			if(range.e.c < C) range.e.c = C;
			var cell = {v: data[R][C] };
			if(cell.v == null) continue;
			var cell_ref = XLSX.utils.encode_cell({c:C,r:R});
			
			if(typeof cell.v === 'number') cell.t = 'n';
			else if(typeof cell.v === 'boolean') cell.t = 'b';
			else if(cell.v instanceof Date) {
				cell.t = 'n'; cell.z = XLSX.SSF._table[14];
				cell.v = datenum(cell.v);
			}
			else cell.t = 's';
			
			ws[cell_ref] = cell;
		}
	}
	if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
	return ws;
}
 
function Workbook() {
	if(!(this instanceof Workbook)) return new Workbook();
	this.SheetNames = [];
	this.Sheets = {};
}
 
function s2ab(s) {
	var buf = new ArrayBuffer(s.length);
	var view = new Uint8Array(buf);
	for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
	return buf;
}

function export_table_to_excel(id) {
var theTable = document.getElementById(id);
var oo = generateArray(theTable);
var ranges = oo[1];
var images = oo[2];

/* original data */
var data = oo[0]; 
var ws_name = "SheetJS";
console.log(data); 

var wb = new Workbook(), ws = sheet_from_array_of_arrays(data);
 
/* add ranges to worksheet */
ws['!merges'] = ranges;

/* add images to the worksheet */
ws['!images'] = images;

/* column formatting properties */
ws['!cols'] = [
    {wpx:150},
    {wpx:150},
    {wpx:150},
    {wpx:150}
];

/* add worksheet to workbook */
wb.SheetNames.push(ws_name);
wb.Sheets[ws_name] = ws;


var wbout = XLSX.write(wb, {bookType:'xlsx', bookSST:false, type: 'binary'});

saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), "test.xlsx")
}