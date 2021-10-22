function main() {
    var accountSelector = MccApp.accounts().withIds(['']);
    var accountIterator = accountSelector.get();

    while (accountIterator.hasNext()) {
        var account = accountIterator.next();
        MccApp.select(account);
        var cuenta = account.getCustomerId();

        var cpalimit = 4;

        var timeZone = AdWordsApp.currentAccount().getTimeZone();
        var date = new Date();
        var fecha = Utilities.formatDate(date, timeZone, "yyyyMMdd");

        var timerange = "LAST_7_DAYS";

        var spreadsheet_url = "";

        //abrimos la hoja de las keywords en la spreadsheet seleccionada

        var ss = SpreadsheetApp.openByUrl(spreadsheet_url);
        var sheet = ss.getSheetByName("KwsCPLLimit");
        var sheetCuotaImpresiones = ss.getSheetByName("CuotaImpresiones")

        if (sheetCuotaImpresiones) {
            sheetCuotaImpresiones.clear();
            var columnasI = ['Keywords ' + timerange, 'Cuota impresiones parte superior > 70', 'Cuota impresiones parte superior < 70', 'Cuota impr. perdidas parte superior > 50'];
            var columnasI_str = columnasI.join(',') + " ";
            sheetCuotaImpresiones.appendRow(columnasI);
        }

        if (sheet) {
            sheet.clear();

            //Títulos de los encabezaos
            var columnAs = ['Fecha',
                'Keyword',
                'Grupo',
                'CPL7',
                'CPL14',
                'CPL30',
                'Conv7',
                'Conv14',
                'Conv30',
                'Cost7',
                'Cost14',
                'Cost30',
            ];
            var columnAs_str = columnAs.join(',') + " ";
            sheet.appendRow(columnAs);
        }

        removeLabelsFromKeywords();
        labels(cpalimit);

    }
}

function removeLabelsFromKeywords() {
    var keywordIterator = AdWordsApp.keywords()
    .forDateRange("ALL_TIME")
    .withCondition("CampaignStatus = ENABLED")
    .withCondition("AdGroupStatus = ENABLED")
    .withCondition("LabelNames CONTAINS_ANY ['CPL-alto-7D', 'CPL-alto-14D', 'CPL-alto-30D', 'CPL-medio-7D', 'CPL-medio-14D', 'CPL-medio-30D', 'CPL-bajo-7D', 'CPL-bajo-14D', 'CPL-bajo-30D', 'Cuota impr. parte sup. alta-30D', 'Cuota impr. parte sup. abs. baja-30D', 'Cuota impr. perdidas parte sup. alta-30D']")
    .get();

    while (keywordIterator.hasNext()) {
        var keyword = keywordIterator.next();
        keyword.removeLabel('CPL-alto-7D');
        keyword.removeLabel('CPL-alto-14D');
        keyword.removeLabel('CPL-alto-30D');
        keyword.removeLabel('CPL-medio-7D');
        keyword.removeLabel('CPL-medio-14D');
        keyword.removeLabel('CPL-medio-30D');
        keyword.removeLabel('CPL-bajo-7D');
        keyword.removeLabel('CPL-bajo-14D');
        keyword.removeLabel('CPL-bajo-30D');
        keyword.removeLabel('Cuota impr. parte sup. alta-30D');
        keyword.removeLabel('Cuota impr. parte sup. abs. baja-30D');
        keyword.removeLabel('Cuota impr. perdidas parte sup. alta-30D');
    }
}


function labels(cpalimit) {

    var keywordIterator = AdWordsApp.keywords()
    .forDateRange("LAST_7_DAYS")
    .withCondition("CampaignStatus = ENABLED")
    .withCondition("AdGroupStatus = ENABLED")
    .withCondition("CampaignName IN ['']")
    .get();


    while (keywordIterator.hasNext()) {
        var keyword = keywordIterator.next();
        var name = "'" + keyword.getText();

        var stats = keyword.getStatsFor("LAST_7_DAYS");
        var Conv = stats.getConversions();  //Número de conversiones
        var Cost = stats.getCost();
        if (Conv > 0) { var cpa = (Cost / Conv); } else { var cpa = 0; }

        var stats14 = keyword.getStatsFor("LAST_14_DAYS");
        var Conv14 = stats14.getConversions();  //Número de conversiones
        var Cost14 = stats14.getCost();

        var cpa14 = (Cost14 / Conv14);
        if (Conv14 > 0) { var cpa14 = (Cost14 / Conv14); } else { var cpa14 = 0; }


        var stats30 = keyword.getStatsFor("LAST_30_DAYS");
        var Conv30 = stats30.getConversions();  //Número de conversiones
        var Cost30 = stats30.getCost();
        var cpa30 = (Cost30 / Conv30);
        if (Conv30 > 0) { var cpa30 = (Cost30 / Conv30); } else { var cpa30 = 0; }


        if (Conv > 0 && cpa > cpalimit) {
            keyword.applyLabel('CPL-alto-7D')
            Logger.log('CPA ALTO ' + name)
        }

        if (Conv > 0 && cpa < cpalimit && cpa > cpalimit / 2) {
            keyword.applyLabel('CPL-medio-7D')
            Logger.log('CPA Medio ' + name)
        }

        if (Conv > 0 && Conv > 1 && cpa < cpalimit / 2) {
            keyword.applyLabel('CPL-bajo-7D')
            Logger.log('CPA Bajo ' + name)
        }

        if (Conv14 > 0 && cpa14 > cpalimit) {
            keyword.applyLabel('CPL-alto-14D')
            Logger.log('CPA ALTO 14 ' + name)
        }

        if (Conv14 > 0 && cpa14 < cpalimit && cpa14 > cpalimit / 2) {
            keyword.applyLabel('CPL-medio-14D')
            Logger.log('CPA Medio 14 ' + name)
        }

        if (Conv14 > 0 && cpa14 < cpalimit / 2) {
            keyword.applyLabel('CPL-bajo-14D')
            Logger.log('CPA Bajo 14 ' + name)
        }

        if (Conv30 > 0 && cpa30 > cpalimit) {
            keyword.applyLabel('CPL-alto-30D')
            Logger.log('CPA ALTO 30' + name)
        }

        if (Conv30 > 0 && cpa30 < cpalimit && cpa30 > cpalimit / 2) {
            keyword.applyLabel('CPL-medio-30D')
            Logger.log('CPA Medio 30 ' + name)
        }

        if (Conv30 > 0 && cpa30 < cpalimit / 2) {
            keyword.applyLabel('CPL-bajo-30D')
            Logger.log('CPA Bajo 30 ' + name)
        }

    }


}

function getImpressionShare(sheet, timerange) {
    //Array definition and get data 

    var keywordIterator = AdsApp.keywords()
        .withCondition("Status = ENABLED")
        .withCondition("AdGroupStatus = ENABLED")
        .withCondition("CampaignStatus = ENABLED")
        .forDateRange(timerange)
        .get();

    var keywordIterator2 = AdsApp.keywords()
        .withCondition("Status = ENABLED")
        .withCondition("AdGroupStatus = ENABLED")
        .withCondition("CampaignStatus = ENABLED")
        .withCondition("LabelNames CONTAINS_ANY ['Cuota impr. parte sup. alta']")
        .forDateRange(timerange)
        .get();

    var keywordIterator3 = AdsApp.keywords()
        .withCondition("Status = ENABLED")
        .withCondition("AdGroupStatus = ENABLED")
        .withCondition("CampaignStatus = ENABLED")
        .withCondition("LabelNames CONTAINS_NONE ['Cuota impr. parte sup. alta']")
        .forDateRange(timerange)
        .get();

    var keywordIterator4 = AdsApp.keywords()
        .withCondition("Status = ENABLED")
        .withCondition("AdGroupStatus = ENABLED")
        .withCondition("CampaignStatus = ENABLED")
        .withCondition("LabelNames CONTAINS_ANY ['Cuota impr. perdidas parte sup. alta']")
        .forDateRange(timerange)
        .get();

    while (keywordIterator.hasNext()) {
        var row_array = [];
        var keyword = keywordIterator.next();
        var name = "'" + keyword.getText();

        row_array.push(name);

        if (keywordIterator2.hasNext()) {
            var keyword2 = keywordIterator2.next();
            var name2 = "'" + keyword2.getText();

            row_array.push(name2);
        }

        if (keywordIterator3.hasNext()) {
            var keyword3 = keywordIterator3.next();
            var name3 = "'" + keyword3.getText();

            row_array.push(name3);
        }

        if (keywordIterator4.hasNext()) {
            var keyword4 = keywordIterator4.next();
            var name4 = "'" + keyword4.getText();

            row_array.push(name4);
        }

        sheet.appendRow(row_array);
    }
}