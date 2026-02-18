/**
 * PPTX File Unzip Service
 *
 * Provides streaming PPTX file extraction using yauzl to avoid memory peaks
 * when processing large files (>10MB).
 *
 * @module services/pptx/unzip
 */

import { promises as fs } from 'fs';
import * as path from 'path';
import yauzl from 'yauzl';
import { logger } from '../../utils/logger';
import { PPTXValidationError } from '../../utils/errors';

/**
 * PPTX internal structure paths
 */
export const PPTX_PATHS = {
  SLIDES: 'ppt/slides/',
  SLIDE_LAYOUTS: 'ppt/slideLayouts/',
  SLIDE_MASTERS: 'ppt/slideMasters/',
  THEME: 'ppt/theme/',
  MEDIA: 'ppt/media/',
  RELATIONSHIPS: 'ppt/slides/_rels/',
  GLOBAL_RELATIONSHIPS: 'ppt/_rels/',
} as const;

/**
 * Extracted PPTX structure
 */
export interface ExtractedPPTX {
  slides: Map<number, string>; // slideIndex -> XML content
  slideLayouts: Map<number, string>; // layoutIndex -> XML content
  slideMasters: Map<number, string>; // masterIndex -> XML content
  themes: Map<string, string>; // themeId -> XML content
  relationships: Map<number, string>; // slideIndex -> relationships XML
  globalRelationships?: string; // global relationships XML
  media: Map<string, Buffer>; // filename -> media data
  metadata: PPTXMetadata;
}

/**
 * PPTX file metadata
 */
export interface PPTXMetadata {
  totalSlides: number;
  totalSlideLayouts: number;
  totalSlideMasters: number;
  totalThemes: number;
  totalMedia: number;
  mediaTypes: Map<string, number>; // extension -> count
  hasEncrypted: boolean;
  hasMacros: boolean;
}

/**
 * Unzip options
 */
export interface UnzipOptions {
  /**
   * Whether to extract media files (default: true)
   */
  extractMedia?: boolean;

  /**
   * Whether to extract themes (default: true)
   */
  extractThemes?: boolean;

  /**
   * Whether to extract slide layouts and masters (default: true)
   */
  extractMasters?: boolean;

  /**
   * Maximum file size in bytes (default: 100MB)
   */
  maxFileSize?: number;

  /**
   * Whether to ignore encryption flag (default: false)
   * Some ZIP creators set encryption flag even when files are not encrypted
   */
  ignoreEncryption?: boolean;
}

/**
 * Default unzip options
 */
const DEFAULT_OPTIONS: Required<UnzipOptions> = {
  extractMedia: true,
  extractThemes: true,
  extractMasters: true,
  maxFileSize: 100 * 1024 * 1024, // 100MB
  ignoreEncryption: false,
};

/**
 * Extract PPTX file to memory using streaming decompression
 *
 * @param pptxPath - Path to PPTX file
 * @param options - Unzip options
 * @returns Promise<ExtractedPPTX> - Extracted PPTX structure
 * @throws {PPTXValidationError} - If PPTX file is invalid or corrupted
 */
export function unzipPPTX(
  pptxPath: string,
  options: UnzipOptions = {}
): Promise<ExtractedPPTX> {
  const opts = { ...DEFAULT_OPTIONS, ...options };

  logger.debug('Starting PPTX extraction', {
    path: pptxPath,
    extractMedia: opts.extractMedia,
    extractThemes: opts.extractThemes,
  });

  return new Promise((resolve, reject) => {
    const extracted: ExtractedPPTX = {
      slides: new Map(),
      slideLayouts: new Map(),
      slideMasters: new Map(),
      themes: new Map(),
      relationships: new Map(),
      media: new Map(),
      metadata: {
        totalSlides: 0,
        totalSlideLayouts: 0,
        totalSlideMasters: 0,
        totalThemes: 0,
        totalMedia: 0,
        mediaTypes: new Map(),
        hasEncrypted: false,
        hasMacros: false,
      },
    };

    // Open zip file with strict validation
    yauzl.open(
      pptxPath,
      {
        strictFileNames: true,
        validateEntrySizes: true,
        lazyEntries: true, // Stream entries instead of loading all into memory
        // decompress option removed - using default decompression
      },
      (err, zipfile) => {
        if (err) {
          logger.error('Failed to open PPTX file', {
            error: err.message,
            path: pptxPath,
          });
          return reject(
            new PPTXValidationError(
              `Failed to open PPTX file: ${err.message}`,
              'INVALID_ZIP'
            )
          );
        }

        if (!zipfile) {
          return reject(
            new PPTXValidationError(
              'Failed to open ZIP file (unknown error)',
              'ZIP_OPEN_FAILED'
            )
          );
        }

        logger.debug('PPTX ZIP file opened successfully');

        // Track entry processing
        let entryCount = 0;
        let processedCount = 0;

        zipfile.on('entry', (entry) => {
          entryCount++;
          const entryPath = entry.fileName;

          // Check for encrypted content
          if (entryPath.includes('Encrypted')) {
            extracted.metadata.hasEncrypted = true;
            logger.warn('Detected encrypted content in PPTX');
          }

          // Check for macros (vbaProject.bin)
          if (entryPath.endsWith('vbaProject.bin')) {
            extracted.metadata.hasMacros = true;
            logger.info('Detected VBA macros in PPTX (will be ignored)');
          }

          // Process different entry types
          if (entryPath.startsWith(PPTX_PATHS.SLIDES) && entryPath.endsWith('.xml')) {
            // Extract slide XML
            extractSlide(zipfile, entry, extracted.slides, opts.ignoreEncryption);
            processedCount++;
          } else if (
            opts.extractMasters &&
            entryPath.startsWith(PPTX_PATHS.SLIDE_LAYOUTS) &&
            entryPath.endsWith('.xml')
          ) {
            // Extract slide layout XML
            extractSlideLayout(zipfile, entry, extracted.slideLayouts);
            processedCount++;
          } else if (
            opts.extractMasters &&
            entryPath.startsWith(PPTX_PATHS.SLIDE_MASTERS) &&
            entryPath.endsWith('.xml')
          ) {
            // Extract slide master XML
            extractSlideMaster(zipfile, entry, extracted.slideMasters);
            processedCount++;
          } else if (
            opts.extractThemes &&
            entryPath.startsWith(PPTX_PATHS.THEME) &&
            entryPath.endsWith('.xml')
          ) {
            // Extract theme XML
            extractTheme(zipfile, entry, extracted.themes);
            processedCount++;
          } else if (
            entryPath.startsWith(PPTX_PATHS.RELATIONSHIPS) &&
            entryPath.endsWith('.xml.rels')
          ) {
            // Extract slide relationships
            extractRelationships(zipfile, entry, extracted.relationships);
            processedCount++;
          } else if (
            entryPath === 'ppt/_rels/presentation.xml.rels'
          ) {
            // Extract global relationships
            extractGlobalRelationships(zipfile, entry, extracted);
            processedCount++;
          } else if (
            opts.extractMedia &&
            entryPath.startsWith(PPTX_PATHS.MEDIA)
          ) {
            // Extract media file
            extractMedia(zipfile, entry, extracted.media, extracted.metadata);
            processedCount++;
          } else {
            // Skip this entry
            zipfile.readEntry();
          }
        });

        zipfile.on('end', () => {
          // Update metadata counts
          extracted.metadata.totalSlides = extracted.slides.size;
          extracted.metadata.totalSlideLayouts = extracted.slideLayouts.size;
          extracted.metadata.totalSlideMasters = extracted.slideMasters.size;
          extracted.metadata.totalThemes = extracted.themes.size;
          extracted.metadata.totalMedia = extracted.media.size;

          logger.info('PPTX extraction completed', {
            totalEntries: entryCount,
            processedEntries: processedCount,
            slides: extracted.metadata.totalSlides,
            media: extracted.metadata.totalMedia,
            hasEncrypted: extracted.metadata.hasEncrypted,
            hasMacros: extracted.metadata.hasMacros,
          });

          resolve(extracted);
        });

        zipfile.on('error', (err) => {
          logger.error('Error during ZIP extraction', {
            error: err.message,
          });
          reject(
            new PPTXValidationError(
              `ZIP extraction error: ${err.message}`,
              'EXTRACTION_ERROR'
            )
          );
        });

        // Start reading entries
        zipfile.readEntry();
      }
    );
  });
}

/**
 * Extract slide XML from ZIP entry
 */
function extractSlide(
  zipfile: yauzl.ZipFile,
  entry: yauzl.Entry,
  slides: Map<number, string>,
  ignoreEncryption: boolean = false
): void {
  const match = entry.fileName.match(/slide(\d+)\.xml$/);
  if (!match) {
    logger.warn(`Invalid slide filename: ${entry.fileName}`);
    zipfile.readEntry();
    return;
  }

  const slideIndex = parseInt(match[1], 10);

  // Check if entry is encrypted and handle accordingly
  const isEncrypted = Boolean(entry.isEncrypted);

  if (isEncrypted && !ignoreEncryption) {
    // Skip encrypted entries when not ignoring encryption
    logger.error(`Slide ${slideIndex} is encrypted, skipping`);
    zipfile.readEntry();
    return;
  }

  // Prepare openReadStream options
  const openOptions: any = {};

  // Note: When isEncrypted is true but ignoreEncryption is true,
  // we try to read anyway (false positive workaround)

  zipfile.openReadStream(entry, openOptions, (err, readStream) => {
    if (err) {
      logger.error(`Failed to open read stream for slide ${slideIndex}`, {
        error: err.message,
        encrypted: entry.isEncrypted,
      });
      zipfile.readEntry();
      return;
    }

    let xml = '';
    readStream.on('data', (chunk) => {
      xml += chunk.toString('utf-8');
    });

    readStream.on('end', () => {
      slides.set(slideIndex, xml);
      logger.debug(`Extracted slide ${slideIndex}`, {
        size: xml.length,
      });
      zipfile.readEntry();
    });

    readStream.on('error', (err) => {
      logger.error(`Error reading slide ${slideIndex}`, {
        error: err.message,
      });
      zipfile.readEntry();
    });
  });
}

/**
 * Extract slide layout XML from ZIP entry
 */
function extractSlideLayout(
  zipfile: yauzl.ZipFile,
  entry: yauzl.Entry,
  layouts: Map<number, string>
): void {
  const match = entry.fileName.match(/slideLayout(\d+)\.xml$/);
  if (!match) {
    zipfile.readEntry();
    return;
  }

  const layoutIndex = parseInt(match[1], 10);

  zipfile.openReadStream(entry, (err, readStream) => {
    if (err) {
      zipfile.readEntry();
      return;
    }

    let xml = '';
    readStream.on('data', (chunk) => {
      xml += chunk.toString('utf-8');
    });

    readStream.on('end', () => {
      layouts.set(layoutIndex, xml);
      zipfile.readEntry();
    });

    readStream.on('error', () => {
      zipfile.readEntry();
    });
  });
}

/**
 * Extract slide master XML from ZIP entry
 */
function extractSlideMaster(
  zipfile: yauzl.ZipFile,
  entry: yauzl.Entry,
  masters: Map<number, string>
): void {
  const match = entry.fileName.match(/slideMaster(\d+)\.xml$/);
  if (!match) {
    zipfile.readEntry();
    return;
  }

  const masterIndex = parseInt(match[1], 10);

  zipfile.openReadStream(entry, (err, readStream) => {
    if (err) {
      zipfile.readEntry();
      return;
    }

    let xml = '';
    readStream.on('data', (chunk) => {
      xml += chunk.toString('utf-8');
    });

    readStream.on('end', () => {
      masters.set(masterIndex, xml);
      zipfile.readEntry();
    });

    readStream.on('error', () => {
      zipfile.readEntry();
    });
  });
}

/**
 * Extract theme XML from ZIP entry
 */
function extractTheme(
  zipfile: yauzl.ZipFile,
  entry: yauzl.Entry,
  themes: Map<string, string>
): void {
  const themeId = path.basename(entry.fileName, '.xml');

  zipfile.openReadStream(entry, (err, readStream) => {
    if (err) {
      zipfile.readEntry();
      return;
    }

    let xml = '';
    readStream.on('data', (chunk) => {
      xml += chunk.toString('utf-8');
    });

    readStream.on('end', () => {
      themes.set(themeId, xml);
      zipfile.readEntry();
    });

    readStream.on('error', () => {
      zipfile.readEntry();
    });
  });
}

/**
 * Extract slide relationships XML from ZIP entry
 */
function extractRelationships(
  zipfile: yauzl.ZipFile,
  entry: yauzl.Entry,
  relationships: Map<number, string>
): void {
  const match = entry.fileName.match(/slide(\d+)\.xml\.rels$/);
  if (!match) {
    zipfile.readEntry();
    return;
  }

  const slideIndex = parseInt(match[1], 10);

  zipfile.openReadStream(entry, (err, readStream) => {
    if (err) {
      zipfile.readEntry();
      return;
    }

    let xml = '';
    readStream.on('data', (chunk) => {
      xml += chunk.toString('utf-8');
    });

    readStream.on('end', () => {
      relationships.set(slideIndex, xml);
      zipfile.readEntry();
    });

    readStream.on('error', () => {
      zipfile.readEntry();
    });
  });
}

/**
 * Extract global relationships XML from ZIP entry
 */
function extractGlobalRelationships(
  zipfile: yauzl.ZipFile,
  entry: yauzl.Entry,
  extracted: ExtractedPPTX
): void {
  zipfile.openReadStream(entry, (err, readStream) => {
    if (err) {
      zipfile.readEntry();
      return;
    }

    let xml = '';
    readStream.on('data', (chunk) => {
      xml += chunk.toString('utf-8');
    });

    readStream.on('end', () => {
      extracted.globalRelationships = xml;
      zipfile.readEntry();
    });

    readStream.on('error', () => {
      zipfile.readEntry();
    });
  });
}

/**
 * Extract media file from ZIP entry
 */
function extractMedia(
  zipfile: yauzl.ZipFile,
  entry: yauzl.Entry,
  media: Map<string, Buffer>,
  metadata: PPTXMetadata
): void {
  const filename = path.basename(entry.fileName);
  const ext = path.extname(filename).toLowerCase();

  zipfile.openReadStream(entry, (err, readStream) => {
    if (err) {
      logger.error(`Failed to open read stream for media ${filename}`, {
        error: err.message,
      });
      zipfile.readEntry();
      return;
    }

    const chunks: Buffer[] = [];

    readStream.on('data', (chunk) => {
      chunks.push(chunk as Buffer);
    });

    readStream.on('end', () => {
      const buffer = Buffer.concat(chunks);
      media.set(filename, buffer);

      // Update media type statistics
      const count = metadata.mediaTypes.get(ext) || 0;
      metadata.mediaTypes.set(ext, count + 1);

      logger.debug(`Extracted media file ${filename}`, {
        size: buffer.length,
        type: ext,
      });

      zipfile.readEntry();
    });

    readStream.on('error', (err) => {
      logger.error(`Error reading media ${filename}`, {
        error: err.message,
      });
      zipfile.readEntry();
    });
  });
}
