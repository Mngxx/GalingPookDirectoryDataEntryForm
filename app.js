if (process.env.NODE_ENV !== "production") {
    require('dotenv').config();
}
const express = require('express');
const path = require('path')
const ejsMate = require('ejs-mate')
const methodOverride = require('method-override');
const catchAsync = require('./utils/catchAsync');
const flash = require('connect-flash');
const session = require('express-session');
const ExpressError = require('./utils/ExpressError');
const { google } = require("googleapis");
const spreadsheetId = process.env.SPREADSHEET_ID;

const auth = new google.auth.GoogleAuth({
    keyFile: "credentials.json",
    scopes: "https://www.googleapis.com/auth/spreadsheets",
});

const app = express();

app.use(express.static(__dirname + '/public'));
app.engine('ejs', ejsMate);

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

app.use(express.urlencoded({ extended: true }));
app.use(methodOverride('_method'))
const secret = "secret"
const sessionConfig = {
    name: 'session',
    secret: secret,
    resave: false,
    saveUninitialized: true,
    cookie: {
        httpOnly: true,
        // secure:true,
        expires: Date.now() + 1000 * 60 * 60 * 24 * 7,
        maxAge: 1000 * 60 * 60 * 24 * 7
    }
}

app.use(session(sessionConfig))
app.use(flash())


function getfullname(fn, mi, ln) {
    var fullname = fn + mi + ln;
    return fullname;
}
function getdata_Id(name, email) {
    var id = name.replace(' ', '') + "-" + email.slice(0, email.indexOf("@"));
    id = id.toLowerCase();
    return id;
}
async function getResource_list() {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: "v4", auth: client });
    var datarows = await googleSheets.spreadsheets.values.batchGet({
        auth,
        spreadsheetId,
        ranges: ["Resource List!A:A", "Resource List!B:B"]
    });
    var enc_list = datarows.data.valueRanges[0].values;
    var categ_list = datarows.data.valueRanges[1].values;
    return { enc_list, categ_list };
}

async function getDataRows_Directory() {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: "v4", auth: client });
    const getRows = await googleSheets.spreadsheets.values.get({
        auth,
        spreadsheetId,
        range: "Directory",
    });
    var Rowdata = getRows.data.values;
    Rowdata.splice(0, 1);
    return Rowdata;
}

const requireLogin = (req, res, next) => {
    if (!req.session.user_id) {
        return res.redirect('/')
    }
    next();
}

app.use((req, res, next) => {
    res.locals.duplicate = req.flash('duplicate');
    res.locals.success = req.flash('success');
    res.locals.error = req.flash('error');
    next();
})

app.get('/', (req, res, next) => {
    res.render('dataentryform/login');
})
app.post('/', catchAsync(async(req, res, next) => {
    const { user, pass } = req.body;
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: "v4", auth: client });
    const getRows = await googleSheets.spreadsheets.values.get({
        auth,
        spreadsheetId,
        range: "Data Entry Form v2",
    });
    var Rowdata = getRows.data.values;
    Rowdata.splice(0, 1);
    if (user == Rowdata[0][0] && pass == Rowdata[0][1]) {
        req.session.user_id = process.env.SPREADSHEET_ID;
        req.flash('success', 'Welcome Back!');
        res.redirect('/entryform');
    }
    else {
        req.flash('error', 'Wrong username or password!');
        res.redirect('/');
    }
}))
app.post('/logout', (req, res) => {
    req.session.user_id = null;
    req.session.destroy();
    res.redirect('/');
})
app.get('/entryform',  catchAsync(async (req, res, next) => {
    const { enc_list, categ_list } = await getResource_list();
    categ_list.splice(0, 1); enc_list.splice(0, 1);
    var data = {
        id: "",
        ln: "",
        fn: "",
        mi: "",
        org: "",
        pos: "",
        email: "",
        cn: "",
        categ: "",
        enc: ""
    }

    res.render('dataentryform/entryform', { data, enc_list, categ_list });
}))
app.put('/entryform', requireLogin, catchAsync(async (req, res, next) => {
    const id = req.body.id;
    const Rowdata = await getDataRows_Directory();
    for (let i = 0; i < Rowdata.length; i++) {
        if (id == Rowdata[i][0]) {
            const { enc_list, categ_list } = await getResource_list();
            categ_list.splice(0, 1); enc_list.splice(0, 1);
            var data = {
                id: Rowdata[i][0],
                ln: Rowdata[i][1],
                fn: Rowdata[i][2],
                mi: Rowdata[i][3],
                org: Rowdata[i][4],
                pos: Rowdata[i][5],
                email: Rowdata[i][6],
                cn: Rowdata[i][7],
                categ: Rowdata[i][8],
                enc: Rowdata[i][9]
            }
            req.flash('success', `Successfully found an entry with an ID of ${id}.`);
            res.render('dataentryform/entryform', { data, enc_list, categ_list });
            return;
        }
    }
    req.flash('error', `There is no ID of ${id} found in the Directory.`);
    res.redirect('/entryform');
}))
app.post('/entryform', requireLogin, catchAsync(async (req, res, next) => {
    var { ln, fn, mi, org, pos, email, cn, categ, enc, consent, savemethod } = req.body;
    if (consent == null) {
        consent = "No";
    }
    var Rownumber = req.body.Rownumber | 0;
    var name = getfullname(fn, mi, ln);
    var unique_id = getdata_Id(name, email);
    var today = new Date();
    var date = today.getFullYear() + '-' + (today.getMonth() + 1) + '-' + today.getDate();
    var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
    var dateTime = date + ' ' + time;
    if (savemethod == "save") {
        const Rowdata = await getDataRows_Directory();
        for (let i = 0; i < Rowdata.length; i++) {
            if (unique_id == Rowdata[i][0]) {
                Rownumber = i + 2;
                const { enc_list, categ_list } = await getResource_list();
                categ_list.splice(0, 1); enc_list.splice(0, 1);
                var data = {
                    id: unique_id,
                    ln: ln,
                    fn: fn,
                    mi: mi,
                    org: org,
                    pos: pos,
                    email: email,
                    cn: cn,
                    categ: categ,
                    enc: enc,
                }
                req.flash('duplicate', 'It seems there is already an existing entry.');
                res.render('dataentryform/entryform', { data, enc_list, categ_list, consent, duplicate: req.flash('duplicate'), Rownumber });
                return;
            }
        }
    }
    else if (savemethod == "overwrite") {
        const client = await auth.getClient();
        const googleSheets = google.sheets({ version: "v4", auth: client });
        await googleSheets.spreadsheets.values.update({
            auth,
            spreadsheetId,
            range: `Directory!${Rownumber}:${Rownumber}`,
            valueInputOption: "USER_ENTERED",
            resource: {
                values: [[unique_id, ln, fn, mi, org, pos, email, cn, categ, enc, consent, dateTime]],
            },
        });
        req.flash('success', `Successfully overwrites the entry with an ID of ${unique_id}.`);
        res.redirect('/entryform');
        return;
    }
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: "v4", auth: client });
    await googleSheets.spreadsheets.values.append({
        auth,
        spreadsheetId,
        range: "Directory",
        valueInputOption: "USER_ENTERED",
        resource: {
            values: [[unique_id, ln, fn, mi, org, pos, email, cn, categ, enc, consent, dateTime]],
        },
    });
    req.flash('success', 'Successfully made a new entry.');
    res.redirect('/entryform');
}))
app.delete('/entryform', requireLogin, catchAsync(async (req, res, next) => {
    const id = req.body.id;
    var Rownumber;
    const Rowdata = await getDataRows_Directory();
    for (let i = 0; i < Rowdata.length; i++) {
        if (id == Rowdata[i][0]) {
            Rownumber = i + 1;
            var batchUpdateRequest = {
                requests: [
                    {
                        deleteDimension: {
                            range: {
                                sheetId: process.env.DIRECTORY_SHEETID,
                                dimension: "ROWS",
                                startIndex: Rownumber,
                                endIndex: Rownumber + 1,
                            }
                        }
                    }
                ]
            }
            const client = await auth.getClient();
            const googleSheets = google.sheets({ version: "v4", auth: client });
            await googleSheets.spreadsheets.batchUpdate({
                auth,
                spreadsheetId,
                resource: batchUpdateRequest
            });
            req.flash('success', `Successfully deleted an entry with an ID of ${id}.`);
            res.redirect('/entryform');
            return;
        }
    }
    req.flash('error', `There is no ID of ${id} found in the Directory.`);
    res.redirect('/entryform');
}))
app.all('*', (req, res, next) => {
    next(new ExpressError('Page Not Found', 404));
})
app.use((err, req, res, next) => {
    const { message = "Something went wrong", statusCode = 500 } = err;
    if (!err.message) err.message = "Oh No!, Something Went Wrong!";
    res.status(statusCode).render('error', { err });
})
const port = process.env.PORT || 3000;
app.listen(port, () => {
    console.log(`Serving on port ${port}`);
})
