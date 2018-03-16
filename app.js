let express = require('express');
let bodyParser = require('body-parser');
let path = require('path');
var expressHbs = require('express-handlebars');
var multer  = require('multer');
var xl = require('excel4node');

var NodeGeocoder = require('node-geocoder');
var fs2 = require('fs');

var options = {
    provider: 'google',   
    // Optional depending on the providers
    httpAdapter: 'https', // Default
    apiKey: '<api-key>',
    formatter: null         // 'gpx', 'string', ...
};

var geocoder = NodeGeocoder(options);

var upload = multer({ dest: 'uploads/' });
var wb = new xl.Workbook();
var ws = wb.addWorksheet('Sheet 1');
var style2 = wb.createStyle({
    font: {
        bold: true,
        color: '#000000',
        size: 14
    }, 
    border: { // §18.8.4 border (Border)
        left: {
            style: 'thin', 
            color: '#000000' // HTML style hex value
        },
        right: {
            style: 'thin',
            color: '#000000'
        },
        top: {
            style: 'thin',
            color: '#000000'
        },
        bottom: {
            style: 'thin',
            color: '#000000'
        }
    }
});
var style = wb.createStyle({
    font: {
        color: '#000000',
        size: 12
    },
    alignment: {
        wrapText: true,
        horizontal: 'left'
    },
    border: { // §18.8.4 border (Border)
        left: {
            style: 'thin', 
            color: '#000000' // HTML style hex value
        },
        right: {
            style: 'thin',
            color: '#000000'
        },
        top: {
            style: 'thin',
            color: '#000000'
        },
        bottom: {
            style: 'thin',
            color: '#000000'
        }
    }
});
ws.cell(1, 1, 1, 5, true).string('Requested Address').style(style2);
ws.cell(1, 6, 1, 11, true).string('Response Address').style(style2);
ws.cell(1, 12, 1, 14, true).string('Lattitude').style(style2);
ws.cell(1, 15, 1, 17, true).string('Longitude').style(style2);
wb.write(__dirname+'/public/'+'Excel.xlsx');

//setting the server
let app = express();

//defining port
const PORT = process.env.PORT || 3000;

// view engine setup
app.engine('.hbs', expressHbs({defaultLayout : 'layout', extname : '.hbs'}));
app.set('view engine', '.hbs');

//setting all required middleware
app.use(express.static(path.join(__dirname, 'public')));
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended : false }));

let row = 1;

app.get('/', (req, res, next) => {
    res.render('index');
});

app.post('/', upload.single('address'), (req, response, next) => {
    if (fs2.existsSync(path.join(__dirname, 'public/geocode.txt'))) {
        fs2.unlinkSync(path.join(__dirname, 'public/geocode.txt'));
    }    
    let addresses = [];
    let geocodes = [];
    var lineReader = require('readline').createInterface({
        input: require('fs').createReadStream(req.file.path)
    });
    lineReader.on('line', function (line) {
        addresses.push(line);        
    });
    lineReader.on('close', function() {
        console.log(`Total addresses are : ${addresses.length}`);
        addresses.forEach((address) => {
            geocoder.geocode(address, function(err, res) {
                row++;                
                if(err) {
                    //writeFile(`Request address |${address}| \nError |${err}|`);
                    //writeFile('----------------------------------------------------------------------------------------------------------');
                    writeToExcel(address, 'error', 'error', 'error', row);
                    console.log(`Request address |${address}| \nError |${err}|`);
                    console.log('----------------------------------------------------------------------------------------------------------')
                    geocodes.push({address : address, lat : 'err', lon : 'err'});
                } else if(res.length > 0){
                    //writeFile(`Request address |${address}| \nResponse address- |${res[0].formattedAddress}| \nlatitude : ${res[0].latitude}\nlongitude : ${res[0].longitude}`);
                   // writeFile('----------------------------------------------------------------------------------------------------------');
                   writeToExcel(address, res[0].formattedAddress, res[0].latitude, res[0].longitude, row); 
                   console.log(`Request address |${address}| \nResponse address- |${res[0].formattedAddress}| \nlatitude : ${res[0].latitude}\nlongitude : ${res[0].longitude}`); 
                    console.log('----------------------------------------------------------------------------------------------------------')
                    //console.log(`response for ${address} with value : lat : ${res[0].latitude}, lon : ${res[0].longitude}`); 
                    //console.log(`response value ${res[0].formattedAddress} with value : lat : ${res[0].latitude}, lon : ${res[0].longitude}`);
                    geocodes.push({address : address, lat : res[0].latitude, lon : res[0].longitude});
                } else {
                    //writeFile(`Request address |${address}| \nResponse address- |${res}|`);
                    //writeFile('----------------------------------------------------------------------------------------------------------');
                    writeToExcel(address, 'No', 'No', 'No', row);
                    console.log(`Request address |${address}| \nResponse address- |${res}|`); 
                    console.log('----------------------------------------------------------------------------------------------------------')
                    geocodes.push({address : address, lat : 'no', lon : 'no'});
                } 
                //console.log(`Total geocodes are : ${geocodes.length}`); 
                if(addresses.length === geocodes.length) {
                    response.render('geocode', { geocodes :  geocodes}); 
                    //wb.write(__dirname+'/public/'+'Excel.xlsx');
                }
            });
        });   
    });    
});

function writeToExcel(requestedAddress, responseAddress, lat, lon, row) {
    ws.cell(row, 1, row, 5, true).string(requestedAddress).style(style);
    ws.cell(row, 6, row, 11, true).string(responseAddress).style(style);
    ws.cell(row, 12, row, 14, true).string(lat+'').style(style);
    ws.cell(row, 15, row, 17, true).string(lon+'').style(style);
    wb.write(__dirname+'/public/'+'Excel.xlsx')
}

function writeFile(data) {
    fs2.appendFile(__dirname+'/public/'+'address.txt', data+'\n', function (err) {
        if (err) throw err;
    });
}

app.listen(PORT, () => {
    console.log('Server running at port '+PORT);
});