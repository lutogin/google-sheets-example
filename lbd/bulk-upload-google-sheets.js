import GoogleSheet from './google-sheets/google-sheet';
import VisenzeService from '../../services/visenze/visenze-service';
import bulkUploadDao from '../bulk-upload-daos/bulk-upload-dao';
import RawWardrobeItemValidator from './bulk-upload-wardrobe-validation';
import ImageService from '../../image/image-service';

import {
  FIELDS_TITLE,
  PROPERTY_NAME,
  PROPERTIES_NAMES,
} from '../../wardrobe-item/wardrobe-item-constants/wardrobe-item-bulk-upload-constants';
import changeCase from '../../../lib/text-case-handler';
import { wardrobeItemDao } from '../../wardrobe-item/wardrobe-item-daos';
import makeWardrobeItemObj from './bulk-upluad-generate-item';

async function getDataFromVisenzeByImgId(imageId) {
  return VisenzeService.recognizeWardrobeItemByImage(imageId);
}

function parseRowFromArray(data, sectionName = null) {
  const titleAccumulator = [];
  if (Array.isArray(data)) {
    for (const value of data) {
      titleAccumulator.push(sectionName ? value[sectionName].title : value.title);
    }
  } else {
    return data ? data.title : '';
  }
  return titleAccumulator.join('\n');
}

function getTitleColumns() {
  return [Object.values(FIELDS_TITLE)];
}

const prepareRawDataCell = data =>
  typeof data === 'string' && data.indexOf('\n') !== -1 ? data.trim().split('\n') : data;

/**
 * Make Prefilled wardrobe items from Visenze request.
 * @param itemsWdi {Object}  Items from Visenze request.
 * @returns {Array}         Prefilled wardrobe items.
 */
async function getPrefilledItems(itemsWdi) {
  const rowsDataArray = [];
  // eslint-disable-next-line guard-for-in,no-restricted-syntax
  for (const imageId in itemsWdi) {
    for (let visenzeItem = 0; visenzeItem < itemsWdi[imageId].length; visenzeItem += 1) {
      const linkToImg = await ImageService.getImageLinkById(imageId);
      rowsDataArray.push([
        `=IMAGE("${linkToImg}"; 1)`,
        `=HYPERLINK("${linkToImg}"; "${imageId}")`,
        '',
        '',
        '',
        parseRowFromArray(
          itemsWdi[imageId][visenzeItem][PROPERTIES_NAMES.CATEGORIES],
          PROPERTY_NAME.CATEGORY,
        ),
        parseRowFromArray(
          itemsWdi[imageId][visenzeItem][PROPERTIES_NAMES.ATTRIBUTES],
          PROPERTY_NAME.ATTRIBUTE,
        ),
        parseRowFromArray(
          itemsWdi[imageId][visenzeItem][PROPERTIES_NAMES.OCCASIONS],
          PROPERTY_NAME.OCCASION,
        ),
        parseRowFromArray(
          itemsWdi[imageId][visenzeItem][PROPERTIES_NAMES.CLOTHES_STYLES],
          PROPERTY_NAME.CLOTHES_STYLE,
        ),
        parseRowFromArray(itemsWdi[imageId][visenzeItem].color),
      ]);
    }
  }

  return rowsDataArray;
}

/**
 * Create google sheets with wardrobe items.
 * @param user {Object}
 * @param itemsWdi {Object}
 * @return {Object}
 */
async function wardrobeItemsToGoogleSheet(user, itemsWdi) {
  const title = `BulkUpload-${user.email}-${new Date(Date.now()).toLocaleString()}`;
  const titlesFields = getTitleColumns();

  const googleSheet = new GoogleSheet();
  const googleSheetData = await googleSheet.createTableBulkUpload(title);

  // Make record in bulk-upload DB.
  await bulkUploadDao.create({
    userId: user.id,
    key: googleSheetData.id,
    link: googleSheetData.link,
  });

  // Append header for sheet.
  await googleSheet.appendRows(1, titlesFields);

  const rowsDataArray = await getPrefilledItems(itemsWdi);

  await googleSheet.appendRows(2, rowsDataArray);
  await googleSheet.fillValidationsRules(rowsDataArray.length);

  return googleSheetData;
}

/**
 * Generate google sheet with wardrobe items from image ID.
 * @param user
 * @param imagesId
 * @return {Promise<{link: *, id: *}>}
 */
async function googleSheetGeneratorFromImg(user, imagesId) {
  const accumulatorVisenzeItems = {};
  const imageDataArr = Array.isArray(imagesId) ? imagesId : [imagesId];
  for (const imageId of imageDataArr) {
    accumulatorVisenzeItems[imageId] = await getDataFromVisenzeByImgId(imageId);
  }

  return wardrobeItemsToGoogleSheet(user, accumulatorVisenzeItems);
}

async function loadGoogleSheetToWardrobe(googleSheetId) {
  const googleSheet = new GoogleSheet(googleSheetId);
  const {
    data: { values },
  } = await googleSheet.getGoogleSheet();
  // Will form an array with objects wardrobe items.
  const header = values[0].slice(0);
  const resultWardrobeArray = [];
  for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    const wardrobeObj = {};
    for (let cellIndex = 1; cellIndex < values[rowIndex].length; cellIndex += 1) {
      // First value is empty for preview, start from 2 cell
      if (values[rowIndex][cellIndex] !== '') {
        // Write only the filled values.
        wardrobeObj[changeCase.stringToCamelCase(header[cellIndex])] = prepareRawDataCell(
          values[rowIndex][cellIndex],
        );
      }
    }
    resultWardrobeArray.push(wardrobeObj);
  }
  const validation = new RawWardrobeItemValidator(resultWardrobeArray);

  // Check validations and generate message if has error.
  const resultValidation = await validation.validateAll();
  const result = { isValid: true, message: [], ids: [] };
  // eslint-disable-next-line guard-for-in,no-restricted-syntax
  for (const validationRow in resultValidation) {
    // eslint-disable-next-line no-restricted-syntax,guard-for-in
    for (const dataValidation in resultValidation[validationRow]) {
      if (!resultValidation[validationRow][dataValidation].isValid) {
        result.isValid = false;
        result.message.push(
          `Row #${validationRow} - ${resultValidation[validationRow][dataValidation].message}`,
        );
      }
    }
  }

  if (result.isValid) {
    // eslint-disable-next-line guard-for-in,no-restricted-syntax
    for (const rawWardrobeItem of resultWardrobeArray) {
      const { id } = await wardrobeItemDao.create(await makeWardrobeItemObj(rawWardrobeItem));
      result.ids.push(id);
    }
  }

  return result;
}

export { googleSheetGeneratorFromImg, loadGoogleSheetToWardrobe };
