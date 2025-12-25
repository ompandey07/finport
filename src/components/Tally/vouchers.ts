// components/Tally/vouchers.ts

export interface TallyVoucher {
  metadata: any;
  [key: string]: any;
}

export interface ColumnMapping {
  [key: string]: number;
}

export interface ExcelRow extends Array<any> {}

/**
 * Formats a date value to Tally's YYYYMMDD format
 */
export const formatTallyDate = (dateValue: any): string => {
  if (!dateValue) return '';
  let date: Date;
  
  if (typeof dateValue === 'number') {
    date = new Date((dateValue - 25569) * 86400 * 1000);
  } else if (typeof dateValue === 'string') {
    date = new Date(dateValue);
  } else if (dateValue instanceof Date) {
    date = dateValue;
  } else {
    return '';
  }

  if (date && !isNaN(date.getTime())) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}${month}${day}`;
  }
  return String(dateValue).replace(/[-\/]/g, '');
};

/**
 * Generates a GUID for Tally vouchers
 */
export const generateGUID = (): string => {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
    const r = Math.random() * 16 | 0;
    const v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
};

/**
 * Auto-maps Excel columns to Tally fields based on pattern matching
 */
export const autoMapColumns = (headers: string[]): ColumnMapping => {
  const patterns: { [key: string]: string[] } = {
    'date': ['date', 'dt', 'voucher date', 'invoice date', 'bill date'],
    'vouchernumber': ['voucher', 'invoice', 'bill', 'number', 'no', 'vch no', 'inv no', 'vch'],
    'partyname': ['party', 'customer', 'vendor', 'supplier', 'buyer', 'name', 'ledger', 'account'],
    'partygstno': ['gst', 'gstin', 'gst no', 'tax no'],
    'stockitemname': ['item', 'stock', 'product', 'goods', 'material', 'description', 'particular'],
    'quantity': ['qty', 'quantity', 'units', 'nos', 'pcs'],
    'unit': ['unit', 'uom', 'measure'],
    'rate': ['rate', 'price', 'unit price', 'mrp', 'per'],
    'amount': ['amount', 'value', 'total', 'net amount', 'net'],
    'godownname': ['godown', 'warehouse', 'location', 'store'],
    'batchname': ['batch', 'lot', 'batch no'],
    'narration': ['narration', 'remarks', 'description', 'notes', 'comment']
  };

  const newMapping: ColumnMapping = {};
  
  Object.keys(patterns).forEach(fieldKey => {
    const fieldPatterns = patterns[fieldKey];
    for (let i = 0; i < headers.length; i++) {
      const headerText = headers[i].toLowerCase();
      for (const pattern of fieldPatterns) {
        if (headerText.includes(pattern) || pattern.includes(headerText)) {
          newMapping[fieldKey] = i;
          return;
        }
      }
    }
  });

  return newMapping;
};

/**
 * Converts Excel data to Tally JSON format
 */
export const convertToTallyJSON = (
  excelData: ExcelRow[],
  columnMapping: ColumnMapping,
  voucherType: string,
  defaultGodown: string,
  salesLedger: string
): { tallymessage: TallyVoucher[] } => {
  const tallyMessages: TallyVoucher[] = [];

  excelData.forEach((row, index) => {
    const getValue = (key: string): any => {
      if (columnMapping[key] !== undefined && columnMapping[key] >= 0) {
        return row[columnMapping[key]] !== undefined ? row[columnMapping[key]] : '';
      }
      return '';
    };

    const date = formatTallyDate(getValue('date'));
    const voucherNumber = String(getValue('vouchernumber') || `VCH-${index + 1}`);
    const partyName = String(getValue('partyname') || 'Cash');
    const stockItemName = String(getValue('stockitemname') || 'Default Item');
    const quantity = parseFloat(getValue('quantity')) || 0;
    const unit = String(getValue('unit') || 'Nos.');
    const rate = parseFloat(getValue('rate')) || 0;
    const amount = parseFloat(getValue('amount')) || (quantity * rate);
    const godownName = String(getValue('godownname') || defaultGodown);
    const batchName = String(getValue('batchname') || 'Primary Batch');
    const partyGSTNo = String(getValue('partygstno') || '');

    const guid = generateGUID();
    const remoteId = generateGUID() + '-' + String(index + 1).padStart(8, '0');

    const voucher: TallyVoucher = {
      "metadata": { "type": "Voucher", "remoteid": remoteId, "vchkey": guid + ":00000008", "vchtype": voucherType, "action": "Create", "objview": "Invoice Voucher View" },
      "oldauditentryids": [{ "metadata": true, "type": "Number" }, "-1"],
      "date": date, "vchstatusdate": date, "guid": guid, "enteredby": "admin", "objectupdateaction": "Alter",
      "vouchertypename": voucherType, "partyname": partyName, "partyledgername": partyName, "vouchernumber": voucherNumber,
      "basicbuyername": partyName, "basicbasepartyname": partyName, "numberingstyle": "Manual",
      "cstformissuetype": "\u0004 Not Applicable", "cstformrecvtype": "\u0004 Not Applicable", "fbtpaymenttype": "Default",
      "persistedview": "Invoice Voucher View", "vchstatustaxadjustment": "Default", "vchstatusvouchertype": voucherType,
      "basicbuyerssalestaxno": partyGSTNo, "basicduedateofpymt": "Cash", "vchgstclass": "\u0004 Not Applicable",
      "vouchertypeorigname": voucherType, "diffactualqty": false, "ismstfromsync": false, "isdeleted": false,
      "issecurityonwhenentered": true, "asoriginal": false, "audited": false, "iscommonparty": false, "forjobcosting": false,
      "isoptional": false, "effectivedate": date, "useforexcise": false, "isforjobworkin": false, "allowconsumption": false,
      "useforinterest": false, "useforgainloss": false, "useforgodowntransfer": false, "useforcompound": false,
      "useforservicetax": false, "isreversechargeapplicable": false, "issystem": false, "isfetchedonly": false,
      "isgstoverridden": false, "iscancelled": false, "isonhold": false, "issummary": false, "isecommercesupply": false,
      "isboenotapplicable": false, "isgstsecsevenapplicable": false, "ignoreeinvvalidation": false,
      "cmpgstisothterritoryassessee": false, "partygstisothterritoryassessee": false, "irnjsonexported": false,
      "irncancelled": false, "ignoregstconflictinmig": false, "isopbaltransaction": false, "ignoregstformatvalidation": false,
      "iseligibleforitc": true, "ignoregstoptionaluncertain": false, "updatesummaryvalues": false, "isewaybillapplicable": false,
      "isdeletedretained": false, "isnull": false, "isexcisevoucher": false, "excisetaxoverride": false,
      "usefortaxunittransfer": false, "isexer1nopoverwrite": false, "isexf2nopoverwrite": false, "isexer3nopoverwrite": false,
      "ignoreposvalidation": false, "exciseopening": false, "useforfinalproduction": false, "istdsoverridden": false,
      "istcsoverridden": false, "istdstcscashvch": false, "includeadvpymtvch": false, "issubworkscontract": false,
      "isvatoverridden": false, "ignoreorigvchdate": false, "isvatpaidatcustoms": false, "isdeclaredtocustoms": false,
      "vatadvancepayment": false, "vatadvpay": false, "iscstdelcaredgoodssales": false, "isvatrestaxinv": false,
      "isservicetaxoverridden": false, "isisdvoucher": false, "isexciseoverridden": false, "isexcisesupplyvch": false,
      "gstnotexported": false, "ignoregstinvalidation": false, "isgstrefund": false, "ovrdnewaybillapplicability": false,
      "isvatprincipalaccount": false, "vchstatusisvchnumused": false, "vchgststatusisincluded": false,
      "vchgststatusisuncertain": false, "vchgststatusisexcluded": false, "vchgststatusisapplicable": false,
      "vchgststatusisgstr2breconciled": false, "vchgststatusisgstr2bonlyinportal": false, "vchgststatusisgstr2bonlyinbooks": false,
      "vchgststatusisgstr2bmismatch": false, "vchgststatusisgstr2bindiffperiod": false, "vchgststatusisreteffdateoverrdn": false,
      "vchgststatusisoverrdn": false, "vchgststatusisstatindiffdate": false, "vchgststatusisretindiffdate": false,
      "vchgststatusmainsectionexcluded": false, "vchgststatusisbranchtransferout": false, "vchgststatusissystemsummary": false,
      "vchstatusisunregisteredrcm": false, "vchstatusisoptional": false, "vchstatusiscancelled": false, "vchstatusisdeleted": false,
      "vchstatusisopeningbalance": false, "vchstatusisfetchedonly": false, "vchgststatusisoptionaluncertain": false,
      "vchstatusisreacceptforhsndone": false, "vchstatusisreaccephsnsixonedone": false, "paymentlinkhasmultiref": false,
      "isshippingwithinstate": false, "isoverseastouristtrans": false, "isdesignatedzoneparty": false, "hascashflow": false,
      "ispostdated": false, "usetrackingnumber": false, "isinvoice": true, "mfgjournal": false, "hasdiscounts": false,
      "aspayslip": false, "iscostcentre": false, "isstxnonrealizedvch": false, "isexcisemanufactureron": false,
      "isblankcheque": false, "isvoid": false, "orderlinestatus": false, "vatisagnstcancsales": false, "vatispurcexempted": false,
      "isvatrestaxinvoice": false, "vatisassesablecalcvch": false, "isvatdutypaid": true, "isdeliverysameasconsignee": false,
      "isdispatchsameasconsignor": false, "isdeletedvchretained": false, "vchonlyaddlinfoupdated": false, "changevchmode": false,
      "resetirnqrcode": false, "alterid": String(12317 + index), "masterid": String(1740 + index),
      "voucherkey": String(197469711368200 + index), "voucherretainkey": String(6957 + index), "vouchernumberseries": "Default",
      "allinventoryentries": [{
        "stockitemname": stockItemName, "isdeemedpositive": false, "isgstassessablevalueoverridden": false,
        "strdisgstapplicable": false, "contentnegispos": false, "islastdeemedpositive": false, "isautonegate": false,
        "iscustomsclearance": false, "istrackcomponent": false, "istrackproduction": false, "isprimaryitem": false,
        "isscrap": false, "rate": `${rate.toFixed(2)}/${unit}`, "amount": amount.toFixed(2),
        "actualqty": ` ${quantity.toFixed(2)} ${unit}`, "billedqty": ` ${quantity.toFixed(2)} ${unit}`,
        "batchallocations": [{
          "godownname": godownName, "batchname": batchName, "indentno": "\u0004 Not Applicable",
          "orderno": "\u0004 Not Applicable", "trackingnumber": "\u0004 Not Applicable", "dynamiccstiscleared": false,
          "amount": amount.toFixed(2), "actualqty": ` ${quantity.toFixed(2)} ${unit}`, "billedqty": ` ${quantity.toFixed(2)} ${unit}`
        }],
        "accountingallocations": [{
          "oldauditentryids": [{ "metadata": true, "type": "Number" }, "-1"], "ledgername": salesLedger,
          "gstclass": "\u0004 Not Applicable", "isdeemedpositive": false, "ledgerfromitem": false, "removezeroentries": false,
          "ispartyledger": false, "gstoverridden": false, "isgstassessablevalueoverridden": false, "strdisgstapplicable": false,
          "strdgstispartyledger": false, "strdgstisdutyledger": false, "contentnegispos": false, "islastdeemedpositive": false,
          "iscapvattaxaltered": false, "iscapvatnotclaimed": false, "amount": amount.toFixed(2)
        }]
      }],
      "ledgerentries": [{
        "oldauditentryids": [{ "metadata": true, "type": "Number" }, "-1"], "ledgername": partyName,
        "gstclass": "\u0004 Not Applicable", "isdeemedpositive": true, "ledgerfromitem": false, "removezeroentries": false,
        "ispartyledger": true, "gstoverridden": false, "isgstassessablevalueoverridden": false, "strdisgstapplicable": false,
        "strdgstispartyledger": false, "strdgstisdutyledger": false, "contentnegispos": false, "islastdeemedpositive": true,
        "iscapvattaxaltered": false, "iscapvatnotclaimed": false, "amount": "-" + amount.toFixed(2)
      }]
    };

    tallyMessages.push(voucher);
  });

  return { "tallymessage": tallyMessages };
};