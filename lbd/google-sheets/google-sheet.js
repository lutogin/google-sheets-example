import { google } from 'googleapis';
import auth from './credentials/credentials-load';
import brandDao from '../../../list/brand/brand-daos/brand-dao';
import colorDao from '../../../list/color/color-daos/color-dao';
import careTypeDao from '../../../list/care-type/care-type-daos/care-type-dao';
import designerDao from '../../../list/designer/designer-daos/designer-dao';
import seasonDao from '../../../list/season/season-daos/season-dao';
import storeDao from '../../../list/store/store-daos/store-dao';
// import chartTypeDao from '../../../size/size-daos/chart-type-dao';
import sizeChartDao from '../../../size/size-chart/size-chart-daos/size-chart-dao';
// import sizeDao from '../../../size/size-daos/size-dao';
import materialDao from '../../../list/material/material-daos/material-dao';
import countryDao from '../../../list/country/country-daos/country-dao';
import itemConditionDao from '../../../list/item-condition/item-condition-daos/item-condition-dao';
import CurrencyService from '../../../list/currency/currency-service';

class GoogleSheet {
  PERMISSION_FOR_SHEETS = {
    type: 'anyone',
    role: 'writer',
  };

  VALIDATION_MAX_ROW_SIZE = 6000; // column "Brand's" = 5333 items.

  MAX_ROW_RANGE = 'A1:AG';

  VALIDATION_DATA_TABLE_TITLE = 'validationData';

  BULK_UPLOAD_DATA_TABLE_TITLE = 'bulkUploadData';

  START_DATA_LINE = 2; // 1st row - header.

  COLOR_COLUMN = 9;

  BRAND_COLUMN = 10;

  YEAR_COLUMN = 11;

  LOCATION_COUNTRY_COLUMN = 12;

  CURRENCY_COLUMN = 14;

  DESIGNER_COLUMN = 17;

  SEASON_COLUMN = 18;

  STORE_PURCHASED_COLUMN = 19;

  // SIZE_TYPE_COLUMN = 20;
  //
  // SIZE_CHART_COLUMN = 21;

  SIZE_COLUMN = 20;

  MATERIAL_COLUMN = 21;

  PURCHASED_DATE_COLUMN = 23;

  COUNTRY_MANUFACTURE_COLUMN = 24;

  CARE_TYPE_COLUMN = 25;

  DUSTBAG_COLUMN = 27;

  RECEIPT_INCLUDED_COLUMN = 28;

  AUTHENTICITY_CARD_COLUMN = 29;

  ITEM_CONDITION_COLUMN = 30;

  VALIDATION_MAPPING = {
    [this.COLOR_COLUMN]: 'A:A',
    [this.BRAND_COLUMN]: 'B:B',
    [this.YEAR_COLUMN]: 'C:C',
    [this.LOCATION_COUNTRY_COLUMN]: 'D:D',
    [this.CURRENCY_COLUMN]: 'E:E',
    [this.CARE_TYPE_COLUMN]: 'F:F',
    [this.DESIGNER_COLUMN]: 'G:G',
    [this.SEASON_COLUMN]: 'H:H',
    [this.STORE_PURCHASED_COLUMN]: 'I:I',
    // [this.SIZE_TYPE_COLUMN]: '',
    // [this.SIZE_CHART_COLUMN]: '',
    [this.SIZE_COLUMN]: 'J:J',
    [this.MATERIAL_COLUMN]: 'K:K',
    [this.COUNTRY_MANUFACTURE_COLUMN]: 'L:L',
    [this.ITEM_CONDITION_COLUMN]: 'M:M',
    [this.DUSTBAG_COLUMN]: 'N:N',
    [this.RECEIPT_INCLUDED_COLUMN]: 'O:O',
    [this.AUTHENTICITY_CARD_COLUMN]: 'P:P',
  };

  DATA_OPTION_INSERT_ROWS = 'INSERT_ROWS';

  DATA_OPTION_OVERWRITE = 'OVERWRITE';

  VALUE_INPUT_OPTION_USER_ENTERED = 'USER_ENTERED';

  VALUE_INPUT_OPTION_RAW = 'RAW';

  constructor(spreadsheetId = null) {
    this.sheetsObject = google.sheets({ version: 'v4', auth });
    this.driveObject = google.drive({ version: 'v3', auth });
    this.spreadsheetId = spreadsheetId || null;
    this.spreadsheetUrl = null;
    this.wardrobeDataTableId = null;
  }

  /**
   * Parse title from data base data.
   * @return {Array}
   */
  initHelper = item => item.title;

  /**
   * Init validation values method.
   */
  async initValuesForValidation() {
    const brandData = await brandDao.getAll();
    const colorData = await colorDao.getAll();
    const careTypeData = await careTypeDao.getAll();
    const designerData = await designerDao.getAll();
    const seasonData = await seasonDao.getAll();
    const storeData = await storeDao.getAll();
    const materialData = await materialDao.getAll();
    const countriesData = await countryDao.getAll();
    const itemConditionData = await itemConditionDao.getAll();

    // Size
    const sizeCharts = await sizeChartDao.get();
    const sizes = [];
    sizeCharts.forEach(sizeChart =>
      sizeChart.sizes.forEach(item => {
        sizes.push(`${sizeChart.title} (${sizeChart.chartType.title}) > ${item.title}`);
      }),
    );

    const baseQuestionData = ['yes', 'no'];

    // Year
    const now = new Date();
    const years = [];

    for (let i = 1970; i <= now.getFullYear(); i += 1) {
      years.push(i);
    }

    const countriesList = [
      'Afghanistan',
      'Åland Islands',
      'Albania',
      'Algeria',
      'American Samoa',
      'Andorra',
      'Angola',
      'Anguilla',
      'Antarctica',
      'Antigua and Barbuda',
      'Argentina',
      'Armenia',
      'Aruba',
      'Australia',
      'Austria',
      'Azerbaijan',
      'Bahamas',
      'Bahrain',
      'Bangladesh',
      'Barbados',
      'Belarus',
      'Belgium',
      'Belize',
      'Benin',
      'Bermuda',
      'Bhutan',
      'Bolivia',
      'Bosnia and Herzegovina',
      'Botswana',
      'Bouvet Island',
      'Brazil',
      'British Indian Ocean Territory',
      'Brunei Darussalam',
      'Bulgaria',
      'Burkina Faso',
      'Burundi',
      'Cambodia',
      'Cameroon',
      'Canada',
      'Cape Verde',
      'Cayman Islands',
      'Central African Republic',
      'Chad',
      'Chile',
      'China',
      'Christmas Island',
      'Cocos (Keeling) Islands',
      'Colombia',
      'Comoros',
      'Congo',
      'Congo, Democratic Republic',
      'Cook Islands',
      'Costa Rica',
      "Cote D'Ivoire",
      'Croatia',
      'Cuba',
      'Cyprus',
      'Czech Republic',
      'Denmark',
      'Djibouti',
      'Dominica',
      'Dominican Republic',
      'Ecuador',
      'Egypt',
      'El Salvador',
      'Equatorial Guinea',
      'Eritrea',
      'Estonia',
      'Ethiopia',
      'Falkland Islands (Malvinas)',
      'Faroe Islands',
      'Fiji',
      'Finland',
      'France',
      'French Guiana',
      'French Polynesia',
      'French Southern Territories',
      'Gabon',
      'Gambia',
      'Georgia',
      'Germany',
      'Ghana',
      'Gibraltar',
      'Greece',
      'Greenland',
      'Grenada',
      'Guadeloupe',
      'Guam',
      'Guatemala',
      'Guernsey',
      'Guinea',
      'Guinea-Bissau',
      'Guyana',
      'Haiti',
      'Heard Island and Mcdonald Islands',
      'Holy See (Vatican City State)',
      'Honduras',
      'Hong Kong',
      'Hungary',
      'Iceland',
      'India',
      'Indonesia',
      'Iran',
      'Iraq',
      'Ireland',
      'Isle of Man',
      'Israel',
      'Italy',
      'Jamaica',
      'Japan',
      'Jersey',
      'Jordan',
      'Kazakhstan',
      'Kenya',
      'Kiribati',
      'Korea (North)',
      'Korea (South)',
      'Kosovo',
      'Kuwait',
      'Kyrgyzstan',
      'Laos',
      'Latvia',
      'Lebanon',
      'Lesotho',
      'Liberia',
      'Libyan Arab Jamahiriya',
      'Liechtenstein',
      'Lithuania',
      'Luxembourg',
      'Macao',
      'Macedonia',
      'Madagascar',
      'Malawi',
      'Malaysia',
      'Maldives',
      'Mali',
      'Malta',
      'Marshall Islands',
      'Martinique',
      'Mauritania',
      'Mauritius',
      'Mayotte',
      'Mexico',
      'Micronesia',
      'Moldova',
      'Monaco',
      'Mongolia',
      'Montserrat',
      'Morocco',
      'Mozambique',
      'Myanmar',
      'Namibia',
      'Nauru',
      'Nepal',
      'Netherlands',
      'Netherlands Antilles',
      'New Caledonia',
      'New Zealand',
      'Nicaragua',
      'Niger',
      'Nigeria',
      'Niue',
      'Norfolk Island',
      'Northern Mariana Islands',
      'Norway',
      'Oman',
      'Pakistan',
      'Palau',
      'Palestinian Territory, Occupied',
      'Panama',
      'Papua New Guinea',
      'Paraguay',
      'Peru',
      'Philippines',
      'Pitcairn',
      'Poland',
      'Portugal',
      'Puerto Rico',
      'Qatar',
      'Reunion',
      'Romania',
      'Russian Federation',
      'Rwanda',
      'Saint Helena',
      'Saint Kitts and Nevis',
      'Saint Lucia',
      'Saint Pierre and Miquelon',
      'Saint Vincent and the Grenadines',
      'Samoa',
      'San Marino',
      'Sao Tome and Principe',
      'Saudi Arabia',
      'Senegal',
      'Serbia',
      'Montenegro',
      'Seychelles',
      'Sierra Leone',
      'Singapore',
      'Slovakia',
      'Slovenia',
      'Solomon Islands',
      'Somalia',
      'South Africa',
      'South Georgia and the South Sandwich Islands',
      'Spain',
      'Sri Lanka',
      'Sudan',
      'Suriname',
      'Svalbard and Jan Mayen',
      'Swaziland',
      'Sweden',
      'Switzerland',
      'Syrian Arab Republic',
      'Taiwan, Province of China',
      'Tajikistan',
      'Tanzania',
      'Thailand',
      'Timor-Leste',
      'Togo',
      'Tokelau',
      'Tonga',
      'Trinidad and Tobago',
      'Tunisia',
      'Turkey',
      'Turkmenistan',
      'Turks and Caicos Islands',
      'Tuvalu',
      'Uganda',
      'Ukraine',
      'United Arab Emirates',
      'United Kingdom',
      'United States',
      'United States Minor Outlying Islands',
      'Uruguay',
      'Uzbekistan',
      'Vanuatu',
      'Venezuela',
      'Viet Nam',
      'Virgin Islands, British',
      'Virgin Islands, U.S.',
      'Wallis and Futuna',
      'Western Sahara',
      'Yemen',
      'Zambia',
      'Zimbabwe',
    ];

    this.validationDatas = [
      {
        column: this.COLOR_COLUMN,
        data: colorData.map(this.initHelper),
      },
      {
        column: this.BRAND_COLUMN,
        data: brandData.map(this.initHelper),
      },
      {
        column: this.YEAR_COLUMN,
        data: years,
      },
      {
        column: this.LOCATION_COUNTRY_COLUMN,
        data: countriesList,
      },
      {
        column: this.CURRENCY_COLUMN,
        data: await CurrencyService.getCurrencyList(),
      },
      {
        column: this.CARE_TYPE_COLUMN,
        data: careTypeData.map(this.initHelper),
      },
      {
        column: this.DESIGNER_COLUMN,
        data: designerData.map(this.initHelper),
      },
      {
        column: this.SEASON_COLUMN,
        data: seasonData.map(this.initHelper),
      },
      {
        column: this.STORE_PURCHASED_COLUMN,
        data: storeData.map(this.initHelper),
      },
      // {
      //   column: this.SIZE_TYPE_COLUMN,
      //   data: chartTypeData.map(this.initHelper),
      // },
      // {
      //   column: this.SIZE_CHART_COLUMN,
      //   data: sizeChartData.map(this.initHelper),
      // },
      {
        column: this.SIZE_COLUMN,
        data: sizes,
      },
      {
        column: this.MATERIAL_COLUMN,
        data: materialData.map(this.initHelper),
      },
      {
        column: this.COUNTRY_MANUFACTURE_COLUMN,
        data: countriesData.map(this.initHelper),
      },
      {
        column: this.ITEM_CONDITION_COLUMN,
        data: itemConditionData.map(this.initHelper),
      },
      {
        column: this.DUSTBAG_COLUMN,
        data: baseQuestionData,
      },
      {
        column: this.RECEIPT_INCLUDED_COLUMN,
        data: baseQuestionData,
      },
      {
        column: this.AUTHENTICITY_CARD_COLUMN,
        data: baseQuestionData,
      },
    ];
  }

  /**
   * Get range for bulk load row.
   * @param rowStart {Number}
   * @param rowEnd {Number}
   * @param tableName {String}
   * @returns {string}
   */
  // eslint-disable-next-line class-methods-use-this
  getRangeForRows(rowStart, rowEnd, tableName = '') {
    return `${tableName}!A${rowStart}:AF${rowEnd}`;
  }

  /**
   * Private method for creat table.
   * @param title
   */
  async createTable(title) {
    return this.sheetsObject.spreadsheets.create({
      resource: {
        properties: {
          title,
        },
      },
    });
  }

  /**
   * Method for create bulk upload table with validation data table.
   * @param title       Table name.
   * @returns {Object}  ID, and link to google sheet.
   */
  async createTableBulkUpload(title) {
    const resultCreate = await this.createTable(title);
    this.spreadsheetId = resultCreate.data.spreadsheetId;
    this.spreadsheetUrl = resultCreate.data.spreadsheetUrl;
    this.wardrobeDataTableId = resultCreate.data.sheets[0].properties.sheetId;

    // Add permission for google sheets for write from link.
    await this.driveObject.permissions.create({
      resource: this.PERMISSION_FOR_SHEETS,
      fileId: this.spreadsheetId,
    });

    const resultUpdate = await this.sheetsObject.spreadsheets.batchUpdate({
      spreadsheetId: this.spreadsheetId,
      requestBody: {
        requests: [
          {
            updateSheetProperties: {
              properties: {
                title: this.BULK_UPLOAD_DATA_TABLE_TITLE,
              },
              fields: 'title',
            },
          },
          {
            addSheet: {
              properties: {
                title: this.VALIDATION_DATA_TABLE_TITLE,
                tabColor: {
                  red: 1.0,
                  green: 0.3,
                  blue: 0.4,
                },
              },
            },
          },
        ],
      },
    });

    this.validationDataTableId = resultUpdate.data.replies[1].addSheet.properties.sheetId;

    await this.fillValidationsData();

    return { id: resultCreate.data.spreadsheetId, link: resultCreate.data.spreadsheetUrl };
  }

  /**
   * Append rows.
   * @param rowStart {int}            Number for append row.
   * @param values {Array}            Array values for row.
   * @param range {String|Object}     Array values for row.
   * @param insertOption {String}
   * @param valueOption {String}
   */
  async appendRows(rowStart, values, range = null, insertOption, valueOption) {
    const insertDataOption = insertOption || this.DATA_OPTION_INSERT_ROWS;
    const valueInputOption = valueOption || this.VALUE_INPUT_OPTION_USER_ENTERED;
    // const currentValueArray = Array.isArray(values) ? values : [values];
    return this.sheetsObject.spreadsheets.values.append({
      spreadsheetId: this.spreadsheetId,
      range: range || this.getRangeForRows(rowStart, rowStart + values.length),
      insertDataOption,
      valueInputOption, // @see https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
      resource: {
        values,
      },
    });
  }

  /**
   * Filled validations table.
   */
  async fillValidationsData() {
    await this.initValuesForValidation();
    const currentColumnData = [];

    for (let i = 0; i < this.validationDatas.length; i += 1) {
      currentColumnData.push(
        this.validationDatas[i].data.map(item => {
          const tmpArray = [];
          for (let j = 0; j < i; j += 1) {
            tmpArray.push('');
          }
          tmpArray.push(item);
          return tmpArray;
        }),
      );
    }

    for (const row of currentColumnData.reverse()) {
      await this.appendRows(
        1,
        row,
        `${this.VALIDATION_DATA_TABLE_TITLE}!A1:A${this.VALIDATION_MAX_ROW_SIZE}`,
        this.DATA_OPTION_OVERWRITE,
        this.VALUE_INPUT_OPTION_RAW,
      );
    }
  }

  /**
   * Fill validation rules for wardrobe items data.
   * @param numberOfFields    Count wardrobe items.
   * @returns {true}
   */
  async fillValidationsRules(numberOfFields) {
    for (let i = 0; i < this.validationDatas.length; i += 1) {
      await this.addValidationRule(numberOfFields, this.validationDatas[i].column, [
        {
          userEnteredValue: this.getValidationRangeFromColumnWardrobeItemNumber(
            this.validationDatas[i].column,
          ),
        },
      ]);
    }

    // add validation for Purchased date column
    await this.addValidationRule(numberOfFields, this.PURCHASED_DATE_COLUMN, [], 'DATE_IS_VALID');

    return true;
  }

  /**
   * Generate range validation
   * @param columnNumber
   * @returns {string}
   */
  getValidationRangeFromColumnWardrobeItemNumber(columnNumber) {
    return `=${this.VALIDATION_DATA_TABLE_TITLE}!${this.VALIDATION_MAPPING[columnNumber]}`;
  }

  /**
   * Added rules for some fields.
   * @param numberOfFields {Number}   Count string for fill validation.
   * @param columnIndex {number}      Number of column for validation field.
   * @param values {Array}            Validation array.
   *                                  See https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/other#ConditionValue
   * @param type {String}
   */
  async addValidationRule(numberOfFields, columnIndex, values, type = 'ONE_OF_RANGE') {
    return this.sheetsObject.spreadsheets.batchUpdate({
      spreadsheetId: this.spreadsheetId,
      requestBody: {
        requests: [
          {
            setDataValidation: {
              range: {
                sheetId: this.wardrobeDataTableId,
                startRowIndex: this.START_DATA_LINE - 1,
                endRowIndex: numberOfFields + 1,
                startColumnIndex: columnIndex,
                endColumnIndex: columnIndex + 1,
              },
              rule: {
                condition: {
                  type,
                  values,
                },
                inputMessage: 'Error value. Select from list.',
                strict: true,
                showCustomUi: true,
              },
            },
          },
        ],
      },
    });
  }

  async getGoogleSheet() {
    return this.sheetsObject.spreadsheets.values.get({
      spreadsheetId: this.spreadsheetId,
      range: this.MAX_ROW_RANGE,
    });
  }
}

export default GoogleSheet;
