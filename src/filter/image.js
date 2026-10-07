/**
 * @license Copyright (c) 2003-2023, CKSource Holding sp. z o.o. All rights reserved.
 * Ported from CKEditor 4 (plugins/pastetools/filter/image.js) for the TinyMCE Paste from Word plugin.
 * Licensed under the terms of the GPL, LGPL or MPL licenses. See LICENSE.md.
 */

import { extractGroupContent, getGroups, removeGroups } from './rtf.js';

// Word puts the pasted images in the RTF clipboard as hex data.
// The HTML clipboard points to them with `file://` paths that a page cannot read.

/**
 * @typedef {Object} ImageData
 * @property {string|null} id Unique id of the image in the RTF content.
 * @property {string|null} hex Image bytes as a hex string. `null` for unsupported types.
 * @property {string} type MIME type, or `'unknown'`.
 */

export const SUPPORTED_IMAGE_TYPES = ['image/png', 'image/jpeg', 'image/gif'];

export const RECOGNIZABLE_IMAGE_TYPES = [
  { marker: /\\pngblip/, type: 'image/png' },
  { marker: /\\jpegblip/, type: 'image/jpeg' },
  { marker: /\\emfblip/, type: 'image/emf' },
  { marker: /\\wmetafile\d/, type: 'image/wmf' },
];

export const RECOGNIZABLE_IMAGE_SIGNATURES = [
  { signature: 'ffd8ff', type: 'image/jpeg' },
  { signature: '47494638', type: 'image/gif' },
  { signature: '89504e47', type: 'image/png' },
];

/**
 * Convert a hex string to bytes.
 *
 * @param {string} hexString
 * @returns {Uint8Array}
 */
export function hexToBytes(hexString) {
  const length = Math.floor(hexString.length / 2);
  const bytes = new Uint8Array(length);

  for (let i = 0; i < length; i++) {
    bytes[i] = parseInt(hexString.substr(i * 2, 2), 16);
  }

  return bytes;
}

/**
 * Encode bytes as Base64.
 *
 * @param {Uint8Array|number[]} bytes
 * @returns {string}
 */
export function bytesToBase64(bytes) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  let out = '';

  for (let i = 0; i < bytes.length; i += 3) {
    const b0 = bytes[i];
    const b1 = i + 1 < bytes.length ? bytes[i + 1] : 0;
    const b2 = i + 2 < bytes.length ? bytes[i + 2] : 0;
    const remaining = bytes.length - i;

    out += chars[b0 >> 2];
    out += chars[((b0 & 0x03) << 4) | (b1 >> 4)];
    out += remaining > 1 ? chars[((b1 & 0x0f) << 2) | (b2 >> 6)] : '=';
    out += remaining > 2 ? chars[b2 & 0x3f] : '=';
  }

  return out;
}

/**
 * Get the MIME type of an RTF image group from its control words.
 *
 * @param {string} imageContent
 * @returns {string} `image/*` or `'unknown'`.
 */
export function getImageType(imageContent) {
  const found = RECOGNIZABLE_IMAGE_TYPES.find((test) => test.marker.test(imageContent));
  return found ? found.type : 'unknown';
}

/**
 * Get the MIME type of an image from its first bytes.
 *
 * @param {Uint8Array} bytes
 * @returns {string|null}
 */
export function getImageTypeFromSignature(bytes) {
  const hexSignature = Array.from(bytes.subarray(0, 4), (byte) => byte.toString(16)).join('');
  const matched = RECOGNIZABLE_IMAGE_SIGNATURES.find((test) => hexSignature.indexOf(test.signature) === 0);

  return matched ? matched.type : null;
}

/**
 * Create a Base64 data URL for an image.
 *
 * @param {ImageData} img
 * @returns {string|null} `null` when the type is not supported.
 */
export function createSrcWithBase64(img) {
  if (!SUPPORTED_IMAGE_TYPES.includes(img.type) || !img.hex) {
    return null;
  }

  const data = typeof img.hex === 'string' ? hexToBytes(img.hex) : img.hex;

  return 'data:' + img.type + ';base64,' + bytesToBase64(data);
}

/**
 * Find the images embedded in RTF content.
 *
 * @param {string} rtfContent
 * @returns {ImageData[]}
 */
export function extractFromRtf(rtfContent) {
  const ret = [];

  // Remove headers, footers, non-Word images and drawn objects.
  rtfContent = removeGroups(rtfContent, '(?:(?:header|footer)[lrf]?|nonshppict|shprslt)');
  const wholeImages = getGroups(rtfContent, 'pict');

  if (!wholeImages) {
    return ret;
  }

  const getImageIndex = (id) => {
    // LibreOffice can omit image ids. Such images are always unique.
    if (typeof id !== 'string') {
      return -1;
    }
    return ret.findIndex((image) => image.id === id);
  };

  const getImageId = (image) => {
    const blipUidMatch = image.match(/\\blipuid (\w+)\}/);
    const blipTagMatch = image.match(/\\bliptag(-?\d+)/);

    if (blipUidMatch) {
      return blipUidMatch[1];
    } else if (blipTagMatch) {
      return blipTagMatch[1];
    }

    return null;
  };

  // RTF breaks the content into lines. Images must not contain whitespace.
  const getImageContent = (image) => extractGroupContent(image).replace(/\s/g, '');

  for (const wholeImage of wholeImages) {
    const currentImage = wholeImage.content;
    const imageId = getImageId(currentImage);
    const imageType = getImageType(currentImage);
    const imageDataIndex = getImageIndex(imageId);
    const isAlreadyExtracted = imageDataIndex !== -1 && ret[imageDataIndex].hex;
    // The same image inserted more than once has the same id.
    const isDuplicated = isAlreadyExtracted && ret[imageDataIndex].type === imageType;
    // An image can be duplicated in another format, right after the original.
    const isAlternateFormat =
      isAlreadyExtracted && ret[imageDataIndex].type !== imageType && imageDataIndex === ret.length - 1;
    // WordArt shapes use the \defshp control word.
    const isWordArtShape = currentImage.indexOf('\\defshp') !== -1;
    const isSupportedType = SUPPORTED_IMAGE_TYPES.includes(imageType);
    // Horizontal lines are pictures in RTF but not in HTML.
    const isHorizontalLine = currentImage.indexOf('fHorizRule') !== -1;

    if (isDuplicated) {
      ret.push(ret[imageDataIndex]);
      continue;
    }

    if (isAlternateFormat || isWordArtShape || isHorizontalLine) {
      continue;
    }

    const newImageData = {
      id: imageId,
      hex: isSupportedType ? getImageContent(currentImage) : null,
      type: imageType,
    };

    if (imageDataIndex !== -1) {
      ret.splice(imageDataIndex, 1, newImageData);
    } else {
      ret.push(newImageData);
    }
  }

  return ret;
}

/**
 * Get the `src` of all `<img>` tags in an HTML string.
 *
 * @param {string} html
 * @returns {string[]}
 */
export function extractTagsFromHtml(html) {
  const regexp = /<img[^>]+src="([^"]+)[^>]+/g;
  const ret = [];
  let item;

  while ((item = regexp.exec(html))) {
    ret.push(item[1]);
  }

  return ret;
}

function escapeRegExp(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Replace the `file://` image sources of the HTML with the images found in the RTF content.
 *
 * @param {string} html HTML already cleaned by the Word filter.
 * @param {string} rtf RTF clipboard content.
 * @param {Object} [options]
 * @param {(img: ImageData) => string|null} [options.createSrc] Build the `src` of an image. Defaults to a data URL.
 * @param {(code: string, data: Object) => void} [options.onError] Called when images cannot be matched.
 * @returns {string}
 */
export function embedRtfImages(html, rtf, options = {}) {
  const createSrc = options.createSrc || createSrcWithBase64;
  const onError = options.onError || (() => {});

  if (!rtf) {
    return html;
  }

  const imgTags = extractTagsFromHtml(html);

  if (imgTags.length === 0) {
    return html;
  }

  const hexImages = extractFromRtf(rtf);

  if (hexImages.length === 0) {
    return html;
  }

  const newSrcValues = hexImages.map((img) => createSrc(img));

  if (imgTags.length !== newSrcValues.length) {
    onError('pastefromword-failed-image-extraction', { rtf: hexImages.length, html: imgTags.length });
    return html;
  }

  // The RTF and HTML images are in the same order, so match them by index.
  for (let i = 0; i < imgTags.length; i++) {
    // Only `file://` images are replaced. Online images keep their URL.
    if (imgTags[i].indexOf('file://') !== 0) {
      continue;
    }

    if (!newSrcValues[i]) {
      onError('pastefromword-unsupported-image', { type: hexImages[i].type, index: i });
      continue;
    }

    const imgRegex = new RegExp('(<img [^>]*src=["\']?)' + escapeRegExp(imgTags[i]));
    html = html.replace(imgRegex, '$1' + newSrcValues[i]);
  }

  return html;
}
