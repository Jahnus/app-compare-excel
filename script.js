(function(config) {
    
    var tempo = config.tempo * 60;
    var loader = $('#loader');
    var units = {};
    var units_group = {};
    var $btnprint = $('#btn-print');
    var session = null;
    var sansExportFlag = false;
    var avecExportFlag = false;
    var cptRapport;
    var details;
    var finalTab;
    var donnees;
    var tabErreurs = [];
    var tabCorres = [];
    var tabIntrouvables = [];
    var traites;
    var cptCorres=0;
    var coutErreurs;
    var devise;
    var nbLignes;   
    var lignesReussies;
    var lignesTraitees;
    var lignesErreurs;
    var uniteEnCours;
    var immatEnCours;
    
    loader.hide();
    $('#result').hide();
    $('#errors').hide();
    document.title = config.name;
    $('.top .wrap p').html(config.name);

    // Wialon script loading
    var url = getHtmlVar("baseUrl") || getHtmlVar("hostUrl") || "https://hst-api.wialon.com";
    loadScript(url+"/wsdk/script/wialon.js", initSdk);

    // Get language
    var lang = getHtmlVar("lang") || config.lang;
    if (["en", "ru", "fr"].indexOf(lang) === -1){
	lang = config.lang;
    }

    // Set translation
    $.localise("lang/", {language:lang, complete:ltranslate});
    translate = $.localise.tr;

    // Load datepicker locale
    if (lang !== "en") {
	loadScript("//apps.wialon.com/plugins/wialon/i18n/" + lang + ".js");
    }
    
    events();

    /** FUNCTION FOR WIALON SCRIPT LOADING */
    function loadScript(src, callback) {
	var script = document.createElement("script");

	script.setAttribute("type", "text/javascript");
	script.setAttribute("charset", "UTF-8");
	script.setAttribute("src", src);

	if (callback && typeof callback === "function") {
            script.onload = callback;
	}

	document.getElementsByTagName("head")[0].appendChild(script);
    }

    /** SDK INITIALIZING */
    function initSdk() {
	var url = getHtmlVar("baseUrl");
	if(!url) {
            url = getHtmlVar("hostUrl");
	}
	if(!url) {
            url = 'https://hst-api.wialon.com';
	}

	var params = {
            'authHash' : getHtmlVar("authHash") || getHtmlVar("access_hash"),
            'sid' : getHtmlVar("sid"),
            'token' : getHtmlVar("token") || getHtmlVar("access_token")
	};

	// Session initialize
	session = wialon.core.Session.getInstance();
	session.initSession(url);
        session.loadLibrary("unitEvents");

	smartLogin(params);
    }

    /** AUTHORIZATION */
    function smartLogin(params) {
	var user = getHtmlVar("user") || "";

	if(params.authHash) {
            session.loginAuthHash(params.authHash, function(code) {loginCallback(code, params, 'authHash');});
	} else if (params.sid) {
            session.duplicate(params.sid, user, true, function(code) {loginCallback(code, params, 'sid');});
	} else if (params.token) {
            session.loginToken(params.token, function(code) {loginCallback(code, params, 'token');});
	} else {
            redirectToLoginPage();
	}
    }

    /** LOGIN CALLBACK */
    function loginCallback(code, params, param) {
	if (code) {
            delete params[param];
            smartLogin(params);
	} else {
            var user = session.getCurrUser();
            user.getLocale(function(arg, locale) {
            // Check for users who have never changed the parameters of the metric
            var fd = (locale && locale.fd) ? locale.fd : '%Y-%m-%E_%H:%M:%S';

            var initDatepickerOpt = {
		wd_orig: locale.wd,
		fd: fd
            };

            var regional = $.datepicker.regional[lang];
            if (regional) {
		$.datepicker.setDefaults(regional);

		// Also wialon locale
		wialon.util.DateTime.setLocale(
			regional.dayNames,
			regional.monthNames,
			regional.dayNamesShort,
			regional.monthNamesShort
		);
            }
            initDatepicker(initDatepickerOpt.fd, initDatepickerOpt.wd_orig);
            init();
            });
	}
    }

    /** REDIRECT TO LOGIN PAGE */
    function redirectToLoginPage() {
	var cur = window.location.href;

	// Remove bad parameters from url
	cur = cur.replace(/\&{0,1}(sid|token|authHash|access_hash|access_token)=\w*/g, '');
	cur = cur.replace(/[\?\&]*$/g, '');

	var url = config.homeUrl + '/login.html?client_id=' + config.name + '&lang=' + lang + '&duration=3600&access_type=0x500&redirect_uri=' + encodeURIComponent(cur);

	window.location.href = url;
    }

    /** DATEPICKER INITIALIZING */
    function initDatepicker(setDateFormat, firstDayOrig) {
	var options = {
		template: '<div class="interval-wialon {className}" id="{id}">' +
				'<div class="iw-select">' +
					'<button data-period="0" type="button" class="iw-period-btn period_0">{yesterday}</button>' +
					'<button data-period="1" type="button" class="iw-period-btn period_1">{today}</button>' +
					'<button data-period="2" type="button" class="iw-period-btn period_2">{week}</button>' +
					'<button data-period="3" type="button" class="iw-period-btn period_3">{month}</button>' +
					'<button data-period="4" type="button" class="iw-period-btn period_4">{custom}</button>' +
				'</div>' +
				'<div class="iw-pickers">' +
					'<input type="text" class="iw-from" id="date-from"/> &ndash; <input type="text" class="iw-to" id="date-to"/>' +
					'<button type="button" class="iw-time-btn">{ok}</button>' +
				'</div>' +
				'<div class="iw-labels">' +
					'<a href="#" class="iw-similar-btn past" data-similar="past"></a> ' +
					'<span class="iw-label"></span> ' +
					'<a href="#" class="iw-similar-btn future" data-similar="future"></a>' +
				'</div>' +
			'</div>',
                labels: {
                        yesterday: translate('Hier'),
        		today: translate("Aujourd'hui"),
                	week: translate('Semaine'),
                        month: translate('Mois'),
                	custom: translate('Personaliser'),
                	ok: "OK"
                },
		datepicker: {},
		onInit: function(){
                    $("#ranging-time-wrap").intervalWialon('set', 3);
		},
		onChange: function(data){
                    currentInterval = $("#ranging-time-wrap").intervalWialon('get');
		},
		onAfterClick: function () {
		},
		tzOffset: wialon.util.DateTime.getTimezoneOffset(),
		now: session.getServerTime()
	};

	options.dateFormat = wialon.util.DateTime.convertFormat(setDateFormat.split('_')[0], true);
	options.firstDay = firstDayOrig;

	$("#ranging-time-wrap").intervalWialon(options);
    }
     
    /** EVENTS */
    function events() {
        
        $('#ranging-time-wrap').on('click', '.iw-period-btn, .iw-time-btn', function() {                        
            if($(this).index() !== 4) {
                if (details.length) {
                    var temp = "Souhaitez-vous changer l'intervale des données à afficher ?";
                    if (confirm(temp)) {
                        details = [];
                        loader.show();
                        compareSansExport();
                    }
                }
            }
	});
	
        $('#importFile').on('click', function() {
            document.getElementById('importFile').disabled = true;
            importToTable(); 
        });
        
        $('#filtreVolMin').on('change', function() {
            if (details) {
                if (avecExportFlag) {
                    afficherRapportAvecExport();
                } else {
                    afficherRapportSansExport();
                }
            }
        });
        
        $('#filtreVolMin').keypress(function(e) {
            var keycode = (e.keyCode ? e.keyCode : e.which);
            if (keycode === '13') {
                if (details.length) {
                    if (avecExportFlag) {
                        afficherRapportAvecExport();
                    } else {
                        afficherRapportSansExport();
                    }
                }
            }
        });

	// Print click
	$btnprint.on('click', function() {
            if (details.length) {
                print();
            } else {
                alert("Il n'y a pas de données à imprimer!");
            }            
	});	
    }
    
    /** PRINT TABLE OF IGNITION */
    function print() {
	var resultCode = '';
	var window_;
        var beginCode = '<!DOCTYPE html><html><head><meta charset="utf-8"><link rel="stylesheet" type="text/css" href="css/style.css"/>' +
                		'</head><body>' +
				'<p style="text-align: right; font-size: 12px;">Intervale - ' + $('.iw-label', '#ranging-time-wrap').text() + '</p>' +
				'<h1 style="text-align: center;">Ecarts pleins</h1>' +
				'<div class="content"><div class="wrap" style="width: 100%;" ">';
	
	var endCode = '</div></div></body></html>';
        var tableCode = $('#result').clone().html() + $('#errors').clone().html();

	resultCode = beginCode + tableCode + endCode;

	window_ = window.open('about:blank', 'Print', 'left=300,top=300,right=500,bottom=500,width=1000,height=500');

	window_.document.open();
	window_.document.write(resultCode);
	window_.document.close();

	setTimeout( function() {
            window_.focus();
            window_.print();
            window_.close();
	}, 500 );

	return this;
    }
    
    /** ELEMENTS TRANSLATION */
    function ltranslate() {
	$btnprint.attr('title', translate('Imprimer'));
        $('#importFile').attr('value', translate('Extraire'));
    }

    /** FUNCTION FOR STRING TRANSLATION */
    function translate(txt){
	var result = txt;
	if (typeof TRANSLATIONS !== "undefined" && typeof TRANSLATIONS === "object" && TRANSLATIONS[txt]) {
            result = TRANSLATIONS[txt];
	}
	return result;
    }

    /** GET URL PARAMETERS */
    function getHtmlVar(name) {
	if (!name) {
            return null;
	}
	var pairs = document.location.search.substr(1).split("&");
	for (var i = 0; i < pairs.length; i++) {
            var pair = pairs[i].split("=");
            if (decodeURIComponent(pair[0]) === name) {
		var param = decodeURIComponent(pair[1]);
		param = param.replace(/[?]/g, '');
		return param;
            }
	}
	return null;
    }
        
    function init() {// Execute after login succeed
        // specify what kind of data should be returned
        var res_flags = wialon.item.Item.dataFlag.base | wialon.item.Resource.dataFlag.reports;
        var unit_flags = wialon.util.Number.or(wialon.item.Item.dataFlag.base, wialon.item.Item.dataFlag.customFields, wialon.item.Item.dataFlag.adminFields);

        var sess = wialon.core.Session.getInstance(); // get instance of current Session
        //sess.loadLibrary("unitEvents"); 
        sess.loadLibrary("resourceReports"); // load Reports Library
        sess.loadLibrary("unitEventRegistrar");
        sess.loadLibrary("itemCustomFields");
        sess.updateDataFlags( // load items to current session
            [{type: "type", data: "avl_resource", flags:res_flags , mode: 0}, // 'avl_resource's specification 
             {type: "type", data: "avl_unit_group", flags: unit_flags, mode: 0},
             {type: "type", data: "avl_unit", flags: unit_flags, mode: 0}], // 'avl_unit's specification
            function (code) { // updateDataFlags callback
                            
                var res = sess.getItems("avl_resource"); // get loaded 'avl_resource's items
                if (!res || !res.length){ msg("Resources not found"); return; } // check if resources found
                            
                units_group = sess.getItems("avl_unit_group");
                
                units = sess.getItems("avl_unit");               
                //chercherImmat();
            }
        );        
    }
                
    function importToTable() {
        loader.show();
        finalTab = [];
        var pleins = [];
        var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xlsx|.xls)$/;
        var titre = formatString($("#excelfile").val());
        var json = [];        

        /*Checks whether the file is a valid excel file*/  
        if (regex.test(titre)) {  
            var xlsxflag = false; /*Flag for checking whether excel is .xls format or .xlsx format*/  
            if ($("#excelfile").val().toLowerCase().indexOf(".xlsx") > 0) {  
                xlsxflag = true;  
            } 
            /*Checks whether the browser supports HTML5*/  
            if (typeof (FileReader) !== "undefined") {  
                var reader = new FileReader();  
                reader.onload = function (e) {  
                    var data = e.target.result; 
                    /*Converts the excel data in to object*/  
                    if (xlsxflag) {  
                        var workbook = XLSX.read(data, { type: 'binary' });  
                    }  
                    else {  
                        var workbook = XLS.read(data, { type: 'binary' });  
                    } 
                    /*Gets all the sheetnames of excel in to a variable*/  
                    var sheet_name_list = workbook.SheetNames;
                    /*for (var i=0; i<sheet_name_list.length; i++) {
                        sheet_name_list[i] = formatTitre(sheet_name_list[i]);
                    }*/

                    var cnt = 0; /*This is used for restricting the script to consider only first sheet of excel*/  
                    sheet_name_list.forEach(function (y) { /*Iterate through all sheets*/  
                        /*Convert the cell value to Json*/  
                        if (xlsxflag) {
                            var exceljson = XLSX.utils.sheet_to_json(workbook.Sheets[y]);  
                        }  
                        else {  
                            var exceljson = XLS.utils.sheet_to_row_object_array(workbook.Sheets[y]);  
                        }  
                        if (exceljson.length > 0 && cnt === 0) {  
                            /*BindTable(exceljson, '#exceltable');  
                            cnt++;*/
                            for (var i=0; i<exceljson.length; i++) {
                                var res = new Map();
                                for (var key in exceljson[i]) {
                                    var newkey = formatString(key).toLowerCase();
                                    var donnee = exceljson[i][key];
                                    res.set(newkey, donnee);                                 
                                }
                                json.push(res);
                            }
                            for (var i=0; i<json.length; i++) {
                                if (json[i].get("designation produit")) {
                                    if (~json[i].get("designation produit").toLowerCase().indexOf('gazole') || ~json[i].get("designation produit").toLowerCase().indexOf('diesel') || ~json[i].get("designation produit").toLowerCase().indexOf('sans plomb') || ~json[i].get("designation produit").toLowerCase().indexOf('sans pl') || ~json[i].get("designation produit").toLowerCase().indexOf('e10')) {
                                        pleins.push(json[i]);
                                    }                                 
                                } else if (json[i].get("code produit (libelle)")) {
                                    if (~json[i].get("code produit (libelle)").toLowerCase().indexOf('gazole') || ~json[i].get("code produit (libelle)").toLowerCase().indexOf('diesel') || ~json[i].get("code produit (libelle)").toLowerCase().indexOf('sans plomb') || ~json[i].get("code produit (libelle)").toLowerCase().indexOf('sans pl') || ~json[i].get("designation produit").toLowerCase().indexOf('e10')) {
                                        pleins.push(json[i]);
                                    }
                                } else if (i === 0) {
                                    alert("La structure du fichier n'est pas supportée! \n" + 
                                        "Pour extraire et importer correctement le fichier, \n" + 
                                        "veuillez renommer les colonnes du fichier excel comme suit: \n" + 
                                        "colonne où apparait le nom/l'immatriculation du véhicule: 'immatriculation'; \n" + 
                                        "colonne où apparait la date du plein: 'date'; \n" + 
                                        "colonne où apparait l'heure du plein: 'heure'; \n" + 
                                        "colonne où apparait le produit (gazole, ss pb, parking, péage...): 'designation produit'; \n" + 
                                        "colonne où apparait le lieu du plein: 'lieu enlevement'; \n" + 
                                        "colonne où apparait le volume du plein: 'quantite'; \n" + 
                                        "colonne où apparait le prix du plein: 'montant ttc'; \n" + 
                                        "colonne où apparait la devise du pays (€, $...): 'devise'; \n" + 
                                        "colonne où apparait le kilométrage saisi: 'kilometrage'.\n" + 
                                        "Pour plus de détails ainsi que pour les informations sur les immatriculations, se réferer à la documentation de l'application.");
                                    loader.hide();
                                    break;
                                }
                            }
                            cnt++;
                        }  
                    });
                    if (pleins[0].get("designation produit")) {
                        for (var i=0; i<pleins.length; i++) {
                            if (!pleins[i].get("date")) {
                                afficherAide();
                                loader.hide();
                                break;
                            }
                            var date = pleins[i].get("date");
                            var year = Number(date.split("/")[2]);
                            var month = Number(date.split("/")[1]);
                            var day = Number(date.split("/")[0]);
                            if (!pleins[i].get("heure")) {
                                afficherAide();
                                loader.hide();
                                break;
                            }
                            var heure = pleins[i].get("heure");
                            var hour = Number(heure.split(":")[0]);
                            var minute = Number(heure.split(":")[1]);
                            var unixDate = toTimestamp(year, month, day, hour, minute, "00");
                            var desc = "plein de " + pleins[i].get("designation produit") + ".";
                            var x = 0;
                            var y = 0;
                            if (!pleins[i].get("lieu enlevement")) {
                                afficherAide();
                                loader.hide();
                                break;
                            }
                            var location = pleins[i].get("lieu enlevement");
                            if (!pleins[i].get("quantite")) {
                                afficherAide();
                                loader.hide();
                                break;
                            }
                            var vol = Number(pleins[i].get("quantite"));
                            if (!pleins[i].get("montant ttc")) {
                                afficherAide();
                                loader.hide();
                                break;
                            }
                            var cost = Number(pleins[i].get("montant ttc"));
                            var dev = 0;
                            if (!pleins[i].get("immatriculation")) {
                                afficherAide();
                                loader.hide();
                                break;
                            }
                            var immat = pleins[i].get("immatriculation");
                            immat = immat.replace(/-/gi, "");
                            immat = immat.replace(/ /gi, "");
                            if (!pleins[i].get("kilometrage")) {
                                afficherAide();
                                loader.hide();
                                break;
                            }
                            var kilo = Number(pleins[i].get("kilometrage"));
                            var prod = pleins[i].get("designation produit");
                            devise = pleins[i].get("devise");
                            var ligneTab = {vehi: "", immat: immat, date: unixDate, year: year, month: month, day: day, hour: hour, minute: minute, description: desc, x: x, y: y, location: location, volume: vol, cost: cost, deviation: dev, kilometrage: kilo, produit: prod, devise: devise};
                            finalTab.push(ligneTab);
                        }
                    } else if (pleins[0].get("code produit (libelle)")) {
                        for (var i=0; i<pleins.length; i++) {
                            if (!pleins[i].get("date de la transaction")) {
                                afficherAide();
                                loader.hide();
                                break;
                            }
                            var date = pleins[i].get("date de la transaction");
                            var year = Number(date.split("/")[2]);
                            var month = Number(date.split("/")[1]);
                            var day = Number(date.split("/")[0]);
                            if (!pleins[i].get("heure de la transaction")) {
                                afficherAide();
                                loader.hide();
                                break;
                            }
                            var heure = pleins[i].get("heure de la transaction");
                            var hour = Number(heure.split(":")[0]);
                            var minute = Number(heure.split(":")[1]);                            
                            var unixDate = toTimestamp(year, month, day, hour, minute, "00");
                            var desc = "plein de " + pleins[i].get("code produit (libelle)") + ".";
                            var x = 0;
                            var y = 0;
                            if (!pleins[i].get("code et nom de la station")) {
                                afficherAide();
                                loader.hide();
                                break;
                            }
                            var location = pleins[i].get("code et nom de la station");
                            if (!pleins[i].get("quantite")) {
                                afficherAide();
                                loader.hide();
                                break;
                            }
                            var vol = Number(pleins[i].get("quantite"));
                            if (!pleins[i].get("montant ttc remise/maj incl.")) {
                                afficherAide();
                                loader.hide();
                                break;
                            }
                            var cost = Number(pleins[i].get("montant ttc remise/maj incl."));
                            var dev = 0;
                            if (!pleins[i].get("libelle estampe (immat./nom)")) {
                                afficherAide();
                                loader.hide();
                                break;
                            }
                            var immat = pleins[i].get("libelle estampe (immat./nom)");
                            immat = immat.replace(/-/gi, "");
                            immat = immat.replace(/ /gi, "");
                            if (!pleins[i].get("kilometrage saisi")) {
                                afficherAide();
                                loader.hide();
                                break;
                            }
                            var kilo = Number(pleins[i].get("kilometrage saisi"));
                            var prod = pleins[i].get("code produit (libelle)");
                            devise = pleins[i].get("code devise du pays de la transaction");
                            var ligneTab = {vehi: "", immat: immat, date: unixDate, year: year, month: month, day: day, hour: hour, minute: minute, description: desc, x: x, y: y, location: location, volume: vol, cost: cost, deviation: dev, kilometrage: kilo, produit: prod, devise: devise};
                            finalTab.push(ligneTab);
                        }
                    }                   
                    finalTab.sort(sortByPlate);
                    finalTab = sortByDatePlate(finalTab);
                    for (var i=0; i<finalTab.length; i++) {
                        if (finalTab[i].devise) {
                            devise = finalTab[i].devise;
                            break;
                        }
                    }
                    
                    var nb = finalTab.length;
                    if (nb!==0) {
                        var temp = "Nous avons extrait " + nb + " lignes de plein. Souhaitez-vous les exporter vers la plateforme avant comparaison?";                    
                        if (confirm(temp)) {
                            sansExportFlag = false;
                            avecExportFlag = true;
                            creerCorres();
                        } else {                            
                            sansExportFlag = true;
                            avecExportFlag = false;
                            compareSansExport();                            
                        }
                    } else {
                        var temp3 = "Nous n'avons trouvé aucune ligne de plein. Souhaitez-vous afficher l'aide?";
                        $('#loader').hide();
                        if (confirm(temp3)) {                            
                            afficherAide();
                        }
                    }                     
                }  
                if (xlsxflag) { /*If excel file is .xlsx extension than creates a Array Buffer from excel*/  
                    reader.readAsArrayBuffer($("#excelfile")[0].files[0]);  
                }  
                else {  
                    reader.readAsBinaryString($("#excelfile")[0].files[0]);  
                }  
            }  
            else {  
                $('#loader').hide();
                alert("Désolé! Votre navigateur ne supporte pas HTML5!");  
            }  
        }  
        else {  
            $('#loader').hide();
            alert("Merci de sélectionner un fichier excel valide ! (.xls ou .xlsx)");  
        }  
    }
    
    function preExport() {
        tabCorres.sort(sortByPlate);             
        uniteEnCours = undefined;
        immatEnCours = undefined;
        nbLignes = finalTab.length;   
        lignesReussies = 0;
        lignesTraitees = 0;
        lignesErreurs = 0;
        coutErreurs = 0;
        exportToWialon();
    }

    function exportToWialon() {  
        if (lignesTraitees===0) {            
            immatEnCours = finalTab[lignesTraitees].immat;
            if (!~tabIntrouvables.indexOf(immatEnCours)) {
                for (var j=0; j<tabCorres.length; j++) {
                    if (tabCorres[j].immat === immatEnCours) {
                        finalTab[lignesTraitees].vehi = tabCorres[j].vehi;
                        uniteEnCours = tabCorres[j].unit;
                        enregistrementPlein();
                        break;
                    }
                }
            } else {
                var minute = finalTab[lignesTraitees].minute;
                if (minute<10) {
                    minute = "0" + minute;
                }
                tabErreurs.push("L'immatriculation " + finalTab[lignesTraitees].immat + " n'est pas enregistré sur la plateforme. Le plein de " + finalTab[lignesTraitees].produit + " de " + finalTab[lignesTraitees].volume + " L, éffectué le " + finalTab[lignesTraitees].day + "/" + finalTab[lignesTraitees].month + "/" + finalTab[lignesTraitees].year + " à " + finalTab[lignesTraitees].hour + "h" + minute + " pour un coût de " + finalTab[lignesTraitees].cost + " " + finalTab[lignesTraitees].devise + " ne peut donc pas être pris en compte.");
                coutErreurs += finalTab[lignesTraitees].cost;
                lignesErreurs++;
                lignesTraitees++;
                if (lignesTraitees === nbLignes) {
                    compareAvecExport();
                } else {
                    exportToWialon();
                }
            }                                           
        } else {
            if (finalTab[lignesTraitees].immat === immatEnCours) {
                if (uniteEnCours) {
                    finalTab[lignesTraitees].vehi = finalTab[lignesTraitees-1].vehi;
                    enregistrementPlein();
                } else {
                    var minute = finalTab[lignesTraitees].minute;
                    if (minute<10) {
                        minute = "0" + minute;
                    }
                    tabErreurs.push("L'immatriculation " + finalTab[lignesTraitees].immat + " n'est pas enregistré sur la plateforme. Le plein de " + finalTab[lignesTraitees].produit + " de " + finalTab[lignesTraitees].volume + " L, éffectué le " + finalTab[lignesTraitees].day + "/" + finalTab[lignesTraitees].month + "/" + finalTab[lignesTraitees].year + " à " + finalTab[lignesTraitees].hour + "h" + minute + " pour un coût de " + finalTab[lignesTraitees].cost + " " + finalTab[lignesTraitees].devise + " ne peut donc pas être pris en compte.");
                    coutErreurs += finalTab[lignesTraitees].cost;
                    lignesErreurs++;
                    lignesTraitees++;
                    if (lignesTraitees === nbLignes) {
                        compareAvecExport();
                    } else {
                        exportToWialon();
                    }
                }
            } else {
                uniteEnCours = undefined;
                immatEnCours = finalTab[lignesTraitees].immat;
                if (!~tabIntrouvables.indexOf(immatEnCours)) {
                    for (var j=0; j<tabCorres.length; j++) {
                        if (tabCorres[j].immat === immatEnCours) {
                            finalTab[lignesTraitees].vehi = tabCorres[j].vehi;
                            uniteEnCours = tabCorres[j].unit;
                            enregistrementPlein();
                            break;
                        }
                    }
                } else {
                    var minute = finalTab[lignesTraitees].minute;
                    if (minute<10) {
                        minute = "0" + minute;
                    }
                    tabErreurs.push("L'immatriculation " + finalTab[lignesTraitees].immat + " n'est pas enregistré sur la plateforme. Le plein de " + finalTab[lignesTraitees].produit + " de " + finalTab[lignesTraitees].volume + " L, éffectué le " + finalTab[lignesTraitees].day + "/" + finalTab[lignesTraitees].month + "/" + finalTab[lignesTraitees].year + " à " + finalTab[lignesTraitees].hour + "h" + minute + " pour un coût de " + finalTab[lignesTraitees].cost + " " + finalTab[lignesTraitees].devise + " ne peut donc pas être pris en compte.");
                    coutErreurs += finalTab[lignesTraitees].cost;
                    lignesErreurs++;
                    lignesTraitees++;
                    if (lignesTraitees === nbLignes) {
                        compareAvecExport();
                    } else {
                        exportToWialon();
                    }
                }                     
            }
        }            
    }        
        
    function enregistrementPlein() {
        uniteEnCours.registryFuelFillingEvent(finalTab[lignesTraitees].date, finalTab[lignesTraitees].description, finalTab[lignesTraitees].x, finalTab[lignesTraitees].y, finalTab[lignesTraitees].location, finalTab[lignesTraitees].volume, finalTab[lignesTraitees].cost, finalTab[lignesTraitees].deviation, function (code) {
            if (code) {
                var minute = finalTab[lignesTraitees].minute;
                if (minute<10) {
                    minute = "0" + minute;
                }
                tabErreurs.push("Le plein de " + finalTab[lignesTraitees].produit + " de " + finalTab[lignesTraitees].volume + " L, éffectué le " + finalTab[lignesTraitees].day + "/" + finalTab[lignesTraitees].month + "/" + finalTab[lignesTraitees].year + " à " + finalTab[lignesTraitees].hour + "h" + minute + ", enregistré pour le véhicule " + finalTab[lignesTraitees].immat + " et pour un coût de " + finalTab[lignesTraitees].cost + " " + finalTab[lignesTraitees].devise + " n'a pas été exporté sur la plateforme pour la raison suivante : " + wialon.core.Errors.getErrorText(code));
                coutErreurs += finalTab[lignesTraitees].cost;
                lignesErreurs++;
                lignesTraitees++;
                if (lignesTraitees === nbLignes) {
                    compareAvecExport();
                } else {
                    exportToWialon();
                }
            } else {
                finalTab[lignesTraitees].export = "ok";
                lignesReussies++;
                lignesTraitees++;
                if (lignesTraitees === nbLignes) {
                    compareAvecExport();
                } else {
                    exportToWialon();
                }
            }
        });
    }
    
    function compareAvecExport() {
        if (lignesReussies !== 0) {
            if (lignesErreurs !== 0) {
                alert(lignesErreurs + " ligne(s) sur " + nbLignes + " n'a(n'ont) pas été exportée(s). \n" + 
                    "Pour extraire et importer correctement le fichier, \n" + 
                    "veuillez renommer les colonnes du fichier excel comme suit: \n" + 
                    "colonne où apparait le nom/l'immatriculation du véhicule: 'immatriculation'; \n" + 
                    "colonne où apparait la date du plein: 'date'; \n" + 
                    "colonne où apparait l'heure du plein: 'heure'; \n" + 
                    "colonne où apparait le produit (gazole, ss pb, parking, péage...): 'designation produit'; \n" + 
                    "colonne où apparait le lieu du plein: 'lieu enlevement'; \n" + 
                    "colonne où apparait le volume du plein: 'quantite'; \n" + 
                    "colonne où apparait le prix du plein: 'montant ttc'; \n" + 
                    "colonne où apparait la devise du pays (€, $...): 'devise'; \n" + 
                    "colonne où apparait le kilométrage saisi: 'kilometrage'.\n" + 
                    "Pour plus de détails ainsi que pour les informations sur les immatriculations, se réferer à la documentation de l'application.");
            }
            cptRapport = 0;        
            details = [];
            executeReport();                        
        } else {
            alert("Aucune ligne n'a été exportée avec succès! \n" + 
                "Pour extraire et importer correctement le fichier, \n" + 
                "veuillez renommer les colonnes du fichier excel comme suit: \n" + 
                "colonne où apparait le nom/l'immatriculation du véhicule: 'immatriculation'; \n" + 
                "colonne où apparait la date du plein: 'date'; \n" + 
                "colonne où apparait l'heure du plein: 'heure'; \n" + 
                "colonne où apparait le produit (gazole, ss pb, parking, péage...): 'designation produit'; \n" + 
                "colonne où apparait le lieu du plein: 'lieu enlevement'; \n" + 
                "colonne où apparait le volume du plein: 'quantite'; \n" + 
                "colonne où apparait le prix du plein: 'montant ttc'; \n" + 
                "colonne où apparait la devise du pays (€, $...): 'devise'; \n" + 
                "colonne où apparait le kilométrage saisi: 'kilometrage'.\n" + 
                "Pour plus de détails ainsi que pour les informations sur les immatriculations, se réferer à la documentation de l'application.");
            $('#loader').hide();
            afficherErreurs();
        }
    }   

    function compareSansExport() {
        cptRapport = 0;        
        details = [];
        tabErreurs = [];
        executeReport();
    }
    
    function afficherRapportAvecExport() {
        details = sortByDiff(details);
        details = suppDoublons(details);
        $('#result-tabhead').html("<tr><th>Véhicule</th><th>Immatriculation</th><th>Conducteur</th><th>Date</th><th>Heure</th><th>Emplacement</th><th>Volume Mesuré</th><th>Volume Carte</th><th>Différence</th></tr>");
        var code = "";
        tabCorres.sort(sortByName);    
        var date;        
        var unitEnCours;
        immatEnCours = undefined;
        
        for (var i=0; i<details.length; i++) {
            if (i===0) {
                unitEnCours = details[i].vehi;
                date = details[i].date;
                for (var j=0; j<tabCorres.length; j++) {
                    if (tabCorres[j].vehi === details[i].vehi) {
                        details[i].immat = tabCorres[j].immat;
                        immatEnCours = tabCorres[j].immat;
                        break;
                    }
                }
            } else {
                if (details[i].vehi===unitEnCours) {
                    if (details[i].date === date) {
                        details[i].volumeCarte = details[i-1].volumeCarte;
                        details[i].difference = details[i].volumeMesure - details[i].volumeCarte;
                        details[i].cost = details[i-1].cost;
                        details[i-1].inut = "ok";
                    } else {
                        date = details[i].date;
                    }
                    if (immatEnCours) {
                        details[i].immat = immatEnCours;
                    }
                } else {
                    unitEnCours = details[i].vehi;
                    immatEnCours = undefined;
                    date = details[i].date;
                    for (var j=0; j<tabCorres.length; j++) {
                        if (tabCorres[j].vehi === details[i].vehi) {
                            details[i].immat = tabCorres[j].immat;
                            immatEnCours = tabCorres[j].immat;
                            break;
                        }
                    }
                }
            }
        }
        
        var newTab = [];
        for (var i=0; i<finalTab.length; i++) {
            if (finalTab[i].export) {
                newTab.push(finalTab[i]);
            }
        }
        
        details.sort(sortByPlate);
        details = sortByDatePlate(details);
        newTab.sort(sortByPlate);
        newTab = sortByDatePlate(newTab);
        for (var i=0; i<details.length; i++) {
            if (!details[i].inut) {
                for (var j=0; j<newTab.length; j++) {
                    if (details[i].immat === newTab[j].immat && !newTab[j].traite) {
                        if (details[i].date > newTab[j].date-tempo && details[i].date < newTab[j].date+tempo) {
                            details[i].cost = newTab[j].cost;
                            newTab[j].traite = "ok";
                            break;
                        } else if (details[i].date < newTab[j].date) {
                            break;
                        }
                    } else if (details[i].immat < newTab[j].immat) {
                        break;
                    }
                }
            }
        }
        
        details.sort(sortByName);
        details = sortByDateName(details);
        var volumeMesureTot = 0;
        var volumeCarteTot = 0;
        var difTot = 0;
        var coutTot = 0; 
        for (var i=0; i<details.length; i++) {
            if (!details[i].inut && details[i].volumeCarte !==0 && details[i].volumeMesure !==0 && details[i].difference<=-Number($('#filtreVolMin').val())) {
                var conduc = details[i].conduc;
                if (conduc === "") {
                    conduc = translate('conducteur non renseigné');
                } else if (conduc === "Drivers...") {
                    conduc = translate('conducteurs multiples');
                }  
                var minute = details[i].minute;
                if (minute<10) {
                    minute = "0" + minute;
                }
                var cout = details[i].cost / details[i].volumeCarte * (-details[i].difference);
                volumeMesureTot += details[i].volumeMesure;
                volumeCarteTot += details[i].volumeCarte;
                difTot += (-details[i].difference);
                coutTot += cout;
                if (Number($('#filtreVolMin').val())===0 && details[i].difference!==0) {
                    code += "<tr style='font-weight: bold;'>"
                } else {
                    code += "<tr>"
                }  
                if (cout>=0) {
                    code += ("<td>" + details[i].vehi + "</td><td>" + details[i].immat + "</td><td>" + conduc + "</td><td>" + details[i].day + "/" + details[i].month + "/" + details[i].year + "</td><td>" + details[i].hour + ":" + minute + "</td><td>" + details[i].lieu + "</td><td>" + details[i].volumeMesure + " L</td><td>" + details[i].volumeCarte + " L</td><td>" + (-details[i].difference) + " L</td><td>" + cout.toFixed(2) + " " + (details[i].devise || devise) + "</td></tr>");
                } else {
                    code += ("<td>" + details[i].vehi + "</td><td>" + details[i].immat + "</td><td>" + conduc + "</td><td>" + details[i].day + "/" + details[i].month + "/" + details[i].year + "</td><td>" + details[i].hour + ":" + minute + "</td><td>" + details[i].lieu + "</td><td>" + details[i].volumeMesure + " L</td><td>" + details[i].volumeCarte + " L</td><td>" + (-details[i].difference) + " L</td><td>coût non renseigné</td></tr>");
                }
            } else if (!details[i].inut && !details[i].toErreursTab && details[i].volumeCarte !==0 && details[i].volumeMesure ===0) {
                var minute = details[i].minute;
                if (minute<10) {
                    minute = "0" + minute;
                }
                var cout = details[i].cost;
                if (cout===0) {
                    tabErreurs.push("<b>Le plein de " + details[i].volumeCarte + " L, éffectué le " + details[i].day + "/" + details[i].month + "/" + details[i].year + " à " + details[i].hour + "h" + minute + " et enregistré pour le véhicule " + details[i].immat + " n'a pas été detecté sur la plateforme!</b>");
                } else {
                    coutErreurs += cout;
                    tabErreurs.push("<b>Le plein de " + details[i].volumeCarte + " L, éffectué le " + details[i].day + "/" + details[i].month + "/" + details[i].year + " à " + details[i].hour + "h" + minute + " et enregistré pour le véhicule " + details[i].immat + " n'a pas été detecté sur la plateforme! Le coût s'éleve à " + details[i].cost + " " + (details[i].devise || devise) + ".</b>");
                }
                details[i].toErreursTab = "ok";
            }
        }
        code += ("<tr><td></td><td></td><td></td><td></td><td></td><td></td><td>" + volumeMesureTot + " L</td><td>" + volumeCarteTot + " L</td><td>" + difTot + " L</td><td>" + coutTot.toFixed(2) + " " + devise + "</td></tr>");
               
        if (tabErreurs.length) {
            afficherErreurs(tabErreurs);
        }
        $('#result-tabbody').html(code);
        $('#result').show();
        loader.hide();
    }
    
    function creerCorres() {
        cptCorres = 0;
        tabCorres = [];
        immatEnCours = undefined;
        
        for (var i=0; i<finalTab.length; i++) {
            if (i===0) {
                immatEnCours = finalTab[i].immat;
                cptCorres++;
                searchItemByPlate(immatEnCours);
            } else {
                if (finalTab[i].immat !== immatEnCours) {
                    immatEnCours = finalTab[i].immat;
                    cptCorres++;
                    searchItemByPlate(immatEnCours);
                }
            }
        }
        
    }
    
    function preAfficherRapportSansExport() {
        details = sortByDiff(details);
        details = suppDoublons(details);
        $('#result-tabhead').html("<tr><th>Véhicule</th><th>Immatriculation</th><th>Conducteur</th><th>Date</th><th>Heure</th><th>Emplacement</th><th>Volume Mesuré</th><th>Volume Carte</th><th>Différence</th><th>Coût perte</th></tr>");
        tabCorres.sort(sortByName);        
        traites = [];
        coutErreurs = 0;        
        var unitEnCours;
        var date;
        immatEnCours = undefined;
        
        for (var i=0; i<details.length; i++) {
            if (i===0) {
                unitEnCours = details[i].vehi;
                date = details[i].date;
                for (var j=0; j<tabCorres.length; j++) {
                    if (tabCorres[j].vehi === details[i].vehi) {
                        details[i].immat = tabCorres[j].immat;
                        immatEnCours = tabCorres[j].immat;
                        break;
                    }
                }
            } else {
                if (details[i].vehi===unitEnCours) {
                    if (details[i].date === date) {
                        details[i].volumeCarte = details[i-1].volumeCarte;
                        details[i].difference = details[i].volumeMesure - details[i].volumeCarte;
                        details[i-1].inut = "ok";
                    } else {
                        date = details[i].date;
                    }
                    if (immatEnCours) {
                        details[i].immat = immatEnCours;
                    }
                } else {
                    unitEnCours = details[i].vehi;
                    immatEnCours = undefined;
                    date = details[i].date;
                    for (var j=0; j<tabCorres.length; j++) {
                        if (tabCorres[j].vehi === details[i].vehi) {
                            details[i].immat = tabCorres[j].immat;
                            immatEnCours = tabCorres[j].immat;
                            break;
                        }
                    }
                }
            }
        }
        
        details.sort(sortByPlate);
        details = sortByDatePlate(details);
        for (var i=0; i<details.length; i++) {
            for (var j=0; j<finalTab.length; j++) {
                if (details[i].immat === finalTab[j].immat && !finalTab[j].verif) {
                    if (details[i].date > finalTab[j].date-tempo && details[i].date < finalTab[j].date+tempo && details[i].volumeMesure !== 0) {
                        if (!devise) {
                            devise = finalTab[j].devise;
                        }
                        details[i].volumeCarte = finalTab[j].volume;
                        details[i].cost = finalTab[j].cost;
                        details[i].difference = details[i].volumeMesure - details[i].volumeCarte;
                        details[i].devise = finalTab[j].devise;
                        finalTab[j].verif = "ok";
                        details[i].verif = "ok";
                        break;
                    } else if (details[i].date<finalTab[j.date]) {
                        break;
                    }
                } else if (details[i].plate<finalTab[j].plate) {
                    break;
                }
            }
        }
        
        for (var i=0; i<finalTab.length; i++) {
            if (!finalTab[i].verif) {
                var minute = finalTab[i].minute;
                if (minute<10) {
                    minute = "0" + minute;
                }
                if (~tabIntrouvables.indexOf(finalTab[i].immat)) {
                    var message = "Le véhicule " + finalTab[i].immat + " n'est pas enregistré sur la plateforme. Le plein de " + finalTab[i].produit + " de " + finalTab[i].volume + " L, éffectué le " + finalTab[i].day + "/" + finalTab[i].month + "/" + finalTab[i].year + " à " + finalTab[i].hour + "h" + minute + " pour un coût de " + finalTab[i].cost + " " + finalTab[i].devise + " ne peut donc pas être pris en compte.";
                    coutErreurs += finalTab[i].cost;
                } else {
                    var message = "<b>Le plein de " + finalTab[i].produit + " de " + finalTab[i].volume + " L, éffectué le " + finalTab[i].day + "/" + finalTab[i].month + "/" + finalTab[i].year + " à " + finalTab[i].hour + "h" + minute + " et enregistré pour le véhicule " + finalTab[i].immat + " n'a pas été detecté sur la plateforme! Le coût s'éleve à " + finalTab[i].cost + " " + finalTab[i].devise + ".</b>";
                    coutErreurs += finalTab[i].cost;
                }
                finalTab[i].verif = "ok";
                traites.push(finalTab[i].date + finalTab[i].immat);
                tabErreurs.push(message);
            }
        } 
        afficherRapportSansExport();
    }
    
    function afficherRapportSansExport() {
        var volumeMesureTot = 0;
        var volumeCarteTot = 0;
        var difTot = 0;
        var coutTot = 0; 
        var code = "";       
        for (var i=0; i<details.length; i++) {
            if (details[i].volumeCarte !== 0 && details[i].volumeMesure !== 0 && details[i].difference<=-Number($('#filtreVolMin').val())) {
                var conduc = details[i].conduc;
                if (conduc === "") {
                    conduc = translate('conducteur non renseigné');
                } else if (conduc === "Drivers...") {
                    conduc = translate('conducteurs multiples');
                }  
                var minute = details[i].minute;
                if (minute<10) {
                    minute = "0" + minute;
                }
                var cout = details[i].cost / details[i].volumeCarte * (-details[i].difference);
                volumeMesureTot += details[i].volumeMesure;
                volumeCarteTot += details[i].volumeCarte;
                difTot += (-details[i].difference);
                coutTot += cout;
                traites.push(details[i].date + details[i].immat); 
                if (Number($('#filtreVolMin').val())===0 && details[i].difference!==0) {
                    code += "<tr style='font-weight: bold;'>";
                } else {
                    code += "<tr>";
                }
                code += ("<td>" + details[i].vehi + "</td><td>" + details[i].immat + "</td><td>" + conduc + "</td><td>" + details[i].day + "/" + details[i].month + "/" + details[i].year + "</td><td>" + details[i].hour + ":" + minute + "</td><td>" + details[i].lieu + "</td><td>" + details[i].volumeMesure + " L</td><td>" + details[i].volumeCarte + " L</td><td>" + (-details[i].difference) + " L</td><td>" + cout.toFixed(2) + " " + (details[i].devise || devise) + "</td></tr>");
            } else if (details[i].volumeCarte !== 0 && details[i].volumeMesure === 0 && !details[i].verif && !details[i].inut) {
                if (!~traites.indexOf(details[i].date + details[i].immat)) {
                    var minute = details[i].minute;
                    if (minute<10) {
                        minute = "0" + minute;
                    }
                    tabErreurs.push("<b>Le plein de " + details[i].volumeCarte + " L, éffectué le " + details[i].day + "/" + details[i].month + "/" + details[i].year + " à " + details[i].hour + "h" + minute + " et enregistré pour le véhicule " + details[i].immat + " n'a pas été detecté sur la plateforme!</b>");
                    details[i].verif = "ok";
                    traites.push(details[i].date + details[i].immat);
                }            
            }
        }
        code += ("<tr><td></td><td></td><td></td><td></td><td></td><td></td><td>" + volumeMesureTot + " L</td><td>" + volumeCarteTot + " L</td><td>" + difTot + " L</td><td>" + coutTot.toFixed(2) + " " + devise + "</td></tr>");
        
        if (tabErreurs.length) {
            afficherErreurs(tabErreurs);
        }
        $('#result-tabbody').html(code);
        $('#result').show();
        loader.hide();
    }
         
    function executeReport(){ // execute selected report
        // get data from corresponding fields
        var id_templ=null;
        var id_unit_group = units_group[cptRapport]['_id']; 

        var sess = wialon.core.Session.getInstance(); // get instance of current Session
        var res = sess.getItems("avl_resource"); // get resource by id
        res = res[0];
        var rapport = res.$$user_reports;
        var nb = 1;
        while (!id_templ) {
            if (rapport[nb]) {
                var nom = rapport[nb].n; 
                if (nom === "Rapport App Compare Excel") {
                    id_templ = rapport[nb].id;
                } else nb++;
            } else nb++;                   
        }
                
        var to = currentInterval[1]; // get current server time (end time of report time interval)
        var from = currentInterval[0]; // calculate start time of report
        // specify time interval object
        var interval = { "from": from, "to": to, "flags": wialon.item.MReport.intervalFlag.absolute };
        var template = res.getReport(id_templ); // get report template by id

        res.execReport(template, id_unit_group, 0, interval, // execute selected report
            function(code, data) { // execReport template
                if(code){ msg(wialon.core.Errors.getErrorText(code)); return; } // exit if error code
                if(!data.getTables().length){ // exit if no tables obtained
                    msg("<b>There is no data generated</b>"); return; }
                else showReportResult(data);
        });
    }

    function showReportResult(result){ // show result after report execute  
        var tables = result.getTables(); // get report tables
        if (!tables) return; // exit if no tables                
        for(var i=0; i < tables.length; i++){ // cycle on tables
                                                                     
            // html contains information about one table
            var html = "";
                     
            result.getTableRows(i, 0, tables[i].rows, // get Table rows
                qx.lang.Function.bind( function(html, code, rows) { // getTableRows callback
                    if (code) {msg(wialon.core.Errors.getErrorText(code)); return;} // exit if error code
                        for(var j in rows) { // cycle on table rows
                            if (typeof rows[j].c === "undefined") continue; // skip empty rows                                                  
                        }
                                        
                        if (rows[0].c.length === 7) {
                            var cpt = 0;
                            for (var k=0; k<rows.length; k++) {
                                result.getRowDetail(0, k, function(cod, col) {                                    
                                    if (cod) {
                                        alert(wialon.core.Errors.getErrorText(cod));
                                        cpt++;
                                        if (cpt === rows.length) {
                                            cptRapport++;
                                            if (cptRapport<units_group.length) {
                                                executeReport();
                                            } else {
                                                details.sort(sortByName);
                                                details = sortByDateName(details);
                                                if (avecExportFlag) {                                                    
                                                    afficherRapportAvecExport();
                                                } else if (sansExportFlag) {
                                                    creerCorres();
                                                }
                                            }
                                        }
                                    } else {
                                        if (col[0].i1 !== 0) {
                                            for (var l=0; l<col.length; l++) {
                                                if (col[l].c[1]) {
                                                    var vehicule = col[l].c[0];
                                                    var unixDate = col[l].t2;
                                                    var date = new Date(unixDate*1000);   
                                                    var year = date.getFullYear();
                                                    var month = date.getMonth()+1;
                                                    var day = date.getDate();
                                                    var hour = date.getHours();
                                                    var minute = date.getMinutes();
                                                    var second = date.getSeconds();
                                                    unixDate -= second;
                                                    var lieu = col[l].c[2].t;
                                                    var volMesure = col[l].c[3];
                                                    if (volMesure === "-----") {
                                                        volMesure = 0;
                                                    } else {
                                                        volMesure = Number(col[l].c[3].replace(/ lt/i, ""));
                                                    }                                                            
                                                    var volCarte = col[l].c[4];
                                                    if (volCarte === "-----") {
                                                        volCarte = 0;
                                                    } else {
                                                        volCarte = Number(volCarte.replace(/ lt/i, ""));
                                                    }
                                                    var dif = col[l].c[5];
                                                    if (dif === "-----") {
                                                        dif = 0;
                                                    } else {
                                                        dif = Number(dif.replace(/ lt/i, ""));
                                                    }
                                                    var conducteur = col[l].c[6];
                                                    var ligne = {vehi: vehicule, immat: "", date: unixDate, year: year, month: month, day: day, hour: hour, minute: minute, lieu: lieu, volumeMesure: volMesure, volumeCarte: volCarte, difference: dif, conduc: conducteur, cost: 0, devise: ""};
                                                    details.push(ligne);
                                                }                                                
                                            }
                                            cpt++;
                                            if (cpt === rows.length) {
                                                cptRapport++;
                                                if (cptRapport<units_group.length) {
                                                    executeReport();
                                                } else {
                                                    details.sort(sortByName);
                                                    details = sortByDateName(details);
                                                    if (avecExportFlag) {                                                        
                                                        afficherRapportAvecExport();
                                                    } else if (sansExportFlag) {
                                                        creerCorres();
                                                    }
                                                }
                                            }
                                        } else {
                                            cpt++;
                                            if (cpt === rows.length) {
                                                cptRapport++;
                                                if (cptRapport<units_group.length) {
                                                    executeReport();
                                                } else {
                                                    details.sort(sortByName);
                                                    details = sortByDateName(details);
                                                    if (avecExportFlag) {
                                                        afficherRapportAvecExport();
                                                    } else if (sansExportFlag) {                                                        
                                                        creerCorres();
                                                    }
                                                }
                                            }
                                        }                                        
                                    }
                                });                                
                            }                            
                        }
                }, this, html)
            );
        }                
    }
    
    function afficherErreurs(tab) {
        $('#error-tabbody').empty();        
        var code = "";
        for (var i=0; i<tab.length; i++) {
            code += ("<tr><td>n°" + (i+1) + "</td><td>" + tab[i] + "</td></tr>");
        }
        if (coutErreurs !== 0) {
            code += ("<tr><td>Coût</td><td>" + coutErreurs.toFixed(2) + " " + devise + "</td></tr>");
        }
        $('#error-tabbody').html(code);
        $('#errors').show();
    }
    
    function afficherAide() {
        alert("Pour extraire et importer correctement le fichier, \n" + 
                "veuillez renommer les colonnes du fichier excel comme suit: \n" + 
                "colonne où apparait le nom/l'immatriculation du véhicule: 'immatriculation'; \n" + 
                "colonne où apparait la date du plein: 'date'; \n" + 
                "colonne où apparait l'heure du plein: 'heure'; \n" + 
                "colonne où apparait le produit (gazole, ss pb, parking, péage...): 'designation produit'; \n" + 
                "colonne où apparait le lieu du plein: 'lieu enlevement'; \n" + 
                "colonne où apparait le volume du plein: 'quantite'; \n" + 
                "colonne où apparait le prix du plein: 'montant ttc'; \n" + 
                "colonne où apparait la devise du pays (€, $...): 'devise'; \n" + 
                "colonne où apparait le kilométrage saisi: 'kilometrage'.\n" + 
                "Pour plus de détails ainsi que pour les informations sur les immatriculations, se réferer à la documentation de l'application.");
    }
    
    function formatString(myString) {
            var rules = {
                            a:"àáâãäå",
                            A:"ÀÁÂ",
                            e:"èéêë",
                            E:"ÈÉÊË",
                            i:"ìíîï",
                            I:"ÌÍÎÏ",
                            o:"òóôõöø",
                            O:"ÒÓÔÕÖØ",
                            u:"ùúûü",
                            U:"ÙÚÛÜ",
                            y:"ÿ",
                            c: "ç",
                            C:"Ç",
                            n:"ñ",
                            N:"Ñ"
                            }; 

            function  getJSONKey(key){
                for (var acc in rules){
                    if (rules[acc].indexOf(key)>-1){return acc;}
                }
            }

            function replaceSpec(Texte){
                var regstring="";
                for (var acc in rules){
                    regstring+=rules[acc];
                }
                reg=new RegExp("["+regstring+"]","g" );
                return Texte.replace(reg,function(t){ return getJSONKey(t); });
            }
            return replaceSpec(myString);
    }

    function toTimestamp(year,month,day,hour,minute,second){
        if (month<10) {
            month = "0" + month;
        }
        if (day<10) {
            day = "0" + day;
        }
        if (hour<10) {
            hour = "0" + hour;
        }
        if (minute<10) {
            minute = "0" + minute;
        }
        var strDate = year + "-" + month + "-" + day + "T" + hour + ":" + minute + ":" + second;
        /*var datum = new Date(Date.UTC(year,month-1,day,hour,minute,second));
        return (datum.getTime()/1000);*/
        return Date.parse(strDate)/1000;
    }

    function sortByName(a, b) {
        return sortString(a, b, "vehi");
    }
    
    function sortByPlate(a, b) {
        return sortString(a, b, "immat");
    }
    
    function sortByDateName(tab) {
        var tabInter;
        var newTab = [];
        var vehi;
        for (var i=0; i<tab.length; i++) {
            if (i===0) {
                vehi = tab[i].vehi;
                tabInter = [];
                tabInter.push(tab[i]);
            } else if (i===tab.length-1) {
                if (tab[i].vehi === vehi) {
                    tabInter.push(tab[i]);
                    tabInter.sort(function (a, b) {
                       return a.date - b.date; 
                    });
                    for (var j=0; j<tabInter.length; j++) {
                        newTab.push(tabInter[j]);
                    }
                } else {
                    tabInter.sort(function (a, b) {
                       return a.date - b.date; 
                    });
                    for (var j=0; j<tabInter.length; j++) {
                        newTab.push(tabInter[j]);
                    }
                    newTab.push(tab[i]);
                }
            } else {
                if (tab[i].vehi === vehi) {
                    tabInter.push(tab[i]);
                } else {
                    tabInter.sort(function (a, b) {
                       return a.date - b.date; 
                    });
                    for (var j=0; j<tabInter.length; j++) {
                        newTab.push(tabInter[j]);
                    }
                    vehi = tab[i].vehi;
                    tabInter = [];
                    tabInter.push(tab[i]);
                }
            }
        }
        return newTab;
    }
    
    function sortByDatePlate(tab) {
        var tabInter;
        var newTab = [];
        var immat;
        for (var i=0; i<tab.length; i++) {
            if (i===0) {
                immat = tab[i].immat;
                tabInter = [];
                tabInter.push(tab[i]);
            } else if (i===tab.length-1) {
                if (tab[i].immat === immat) {
                    tabInter.push(tab[i]);
                    tabInter.sort(function (a, b) {
                       return a.date - b.date; 
                    });
                    for (var j=0; j<tabInter.length; j++) {
                        newTab.push(tabInter[j]);
                    }
                } else {
                    tabInter.sort(function (a, b) {
                       return a.date - b.date; 
                    });
                    for (var j=0; j<tabInter.length; j++) {
                        newTab.push(tabInter[j]);
                    }
                    newTab.push(tab[i]);
                }
            } else {
                if (tab[i].immat === immat) {
                    tabInter.push(tab[i]);
                } else {
                    tabInter.sort(function (a, b) {
                       return a.date - b.date; 
                    });
                    for (var j=0; j<tabInter.length; j++) {
                        newTab.push(tabInter[j]);
                    }
                    immat = tab[i].immat;
                    tabInter = [];
                    tabInter.push(tab[i]);
                }
            }
        }
        return newTab;
    }
    
    function sortByDiff(tab) {
        var tabInter;
        var newTab = [];
        var vehi;
        var date;
        for (var i=0; i<tab.length; i++) {
            if (i===0) {
                vehi = tab[i].vehi;
                date = tab[i].date;
                tabInter = [];
                tabInter.push(tab[i]);
            } else if (i===tab.length-1) {
                if (tab[i].vehi===vehi) {
                    if (tab[i].date===date) {
                        tabInter.push(tab[i]);
                        tabInter.sort(function (a, b) {
                            return a.difference - b.difference;
                        });
                        for (var j=0; j<tabInter.length; j++) {
                            newTab.push(tabInter[j]);
                        }
                    } else {
                        tabInter.sort(function (a, b) {
                            return a.difference - b.difference;
                        });
                        for (var j=0; j<tabInter.length; j++) {
                            newTab.push(tabInter[j]);
                        }
                        newTab.push(tab[i]);
                    }                    
                } else {
                    tabInter.sort(function (a, b) {
                        return a.difference - b.difference;
                    });
                    for (var j=0; j<tabInter.length; j++) {
                        newTab.push(tabInter[j]);
                    }
                    newTab.push(tab[i]);
                }
            } else {
                if (tab[i].vehi===vehi) {
                    if (tab[i].date===date) {
                        tabInter.push(tab[i]);
                    } else {
                        tabInter.sort(function (a, b) {
                            return a.difference - b.difference;
                        });
                        for (var j=0; j<tabInter.length; j++) {
                            newTab.push(tabInter[j]);
                        }
                        date = tab[i].date;
                        tabInter = [];
                        tabInter.push(tab[i]);
                    }
                } else {
                    tabInter.sort(function (a, b) {
                        return a.difference - b.difference;
                    });
                    for (var j=0; j<tabInter.length; j++) {
                        newTab.push(tabInter[j]);
                    }
                    vehi = tab[i].vehi;
                    date = tab[i].date;
                    tabInter = [];
                    tabInter.push(tab[i]);
                }
            }
        }
        return newTab;
    }

    function sortString(a, b, key) {
        var v1 = a[key].toLowerCase();
        var v2 = b[key].toLowerCase();
        if (v1<v2) return -1;
        if (v1>v2) return 1;
        return 0;
    }
    
    function chercherImmat() {
        donnees = [];
        for (var i=0; i<units.length; i++) {
            var immat = units[i].getCustomProperty("registration_plate");
            var ligne = {vehi: units[i].getName(), immat: immat};
            donnees.push(ligne);
        }
        /*units[index].getCustomProperty("Registration_plate", "indéfinie", function(rep) {
            donnees.push(rep);
            index++;
            if (index<units.length) {
                chercherImmat();
            }
        });   */     
    }
    
    function searchItemByPlate(immat) {
        var mask1 = immat.toLowerCase();
        var mask2 = immat.toUpperCase();
        var message = "";
        var searchSpec = {
            itemsType:"avl_unit", // Type of the required elements of Wialon
            propName: "rel_profilefield_value", // Name of the characteristic according to which the search will be carried out
            propValueMask: "=" + mask1 + "|" + mask2,   // Meaning of the characteristic: can be used * | , > < =
            sortType: "rel_profilefield_value" // The name of the characteristic according to which you will be sorting a response
        };
	var dataFlags = wialon.item.Item.dataFlag.base;
        
        session.searchItems(searchSpec, true, dataFlags, 0, 0, function(code, data) {            
            if (code) {
                message = "Erreur : " + (wialon.core.Errors.getErrorText(code)) + " pour le véhicule immatriculé : " + immat;
		tabErreurs.push(message);
                if ((tabErreurs.length + tabCorres.length)===cptCorres) {
                    if (sansExportFlag) {
                        preAfficherRapportSansExport();
                    } else {
                        preExport();
                    }                    
                }
            } else {
                if (data.totalItemsCount === 0) {
                    tabIntrouvables.push(immat);                    
                    if ((tabErreurs.length + tabCorres.length + tabIntrouvables.length)===cptCorres) {
                        if (sansExportFlag) {
                            preAfficherRapportSansExport();
                        } else {
                            preExport();
                        } 
                    }
                } else if (data.totalItemsCount === 1) {
                    var unit = data['items'][0];
                    tabCorres.push({vehi: unit.getName(), immat: immat, unit: unit}); 
                    if ((tabErreurs.length + tabCorres.length + tabIntrouvables.length)===cptCorres) {
                        if (sansExportFlag) {
                            preAfficherRapportSansExport();
                        } else {
                            preExport();
                        }
                    }
                } else {
                    message = "L'immatriculation : " + immat + " a été trouvée pour plusieurs véhicules. Veuillez vérifier votre compte sur la plateforme.";
                    tabErreurs.push(message);
                    if ((tabErreurs.length + tabCorres.length + tabIntrouvables.length)===cptCorres) {
                        if (sansExportFlag) {
                            preAfficherRapportSansExport();
                        } else {
                            preExport();
                        }
                    }
                }     
            }                 
	});
    }
    
    function suppDoublons(tab) {
        var tampon = [];
        for (var i=0; i<tab.length; i++) {
            if (i===0) {
                tampon.push(tab[i]);
            } else {
                if (tab[i].vehi!==tab[i-1].vehi) {
                    tampon.push(tab[i]);
                } else {
                    if (tab[i].date!==tab[i-1].date) {
                        tampon.push(tab[i]);
                    } else {
                        if (tab[i].difference!==tab[i-1].difference) {
                            tampon.push(tab[i]);
                        }
                    }
                }
            }
        }
        return tampon;
    }
    
    function msg(text) { alert(text); }
    
})(config);