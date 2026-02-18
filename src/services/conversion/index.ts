/**
 * Main Conversion Service
 *
 * Core service that orchestrates the PPTX to PPTist conversion process.
 * Integrates all conversion components and manages the conversion workflow.
 *
 * @module services/conversion
 */

import { promises as fs } from 'fs';
import * as path from 'path';
import { XMLParser } from 'fast-xml-parser';
import { unzipPPTX, type ExtractedPPTX } from '../pptx/unzip';
import { extractMediaFiles } from '../pptx/extract-media';
import { validatePPTXFile } from '../pptx/validator';
import { ConversionOrchestrator } from './orchestrator.js';
import type { ConversionContext } from '../../types/converters';
import type { ConversionResult } from '../../models/result';
import type { ConversionMetadata } from '../../models/metadata';
import type { ConversionWarning } from '../../models/warning';
import { logger } from '../../utils/logger';
import { PPTXValidationError } from '../../utils/errors';
import { serializeResult } from './serializer.js';
import { collectMetadata } from './collector.js';
import { collectWarnings } from './warnings.js';
import { getDefaultResultsDir, getDefaultMediaDir } from '../../utils/paths';

/**
 * Conversion service options
 */
export interface ConversionServiceOptions {
  /**
   * Task ID for tracking
   */
  taskId: string;

  /**
   * Original filename
   */
  filename: string;

  /**
   * File path to PPTX
   */
  filePath: string;

  /**
   * Output directory for results (default: /tmp/pptx-results/)
   */
  outputDir?: string;

  /**
   * Whether to extract media files (default: true)
   */
  extractMedia?: boolean;

  /**
   * Whether to include animations (default: true)
   */
  includeAnimations?: boolean;

  /**
   * Whether to include notes (default: true)
   */
  includeNotes?: boolean;

  /**
   * PPTist target version (default: 'latest')
   */
  targetVersion?: string;

  /**
   * Skip PPTX validation (useful for testing, default: false)
   */
  skipValidation?: boolean;

  /**
   * Ignore encryption flag (default: false)
   * Some ZIP creators set encryption flag even when files are not encrypted
   */
  ignoreEncryption?: boolean;

  /**
   * Progress callback
   */
  onProgress?: (progress: number, message: string) => void;
}

/**
 * Conversion result
 */
export interface ConversionResultData {
  success: boolean;
  data?: any;
  metadata?: ConversionMetadata;
  warnings?: ConversionWarning[];
  error?: string;
}

/**
 * Main conversion service class
 */
export class ConversionService {
  /**
   * Convert PPTX file to PPTist JSON
   *
   * @param options - Conversion options
   * @returns Conversion result
   */
  async convert(options: ConversionServiceOptions): Promise<ConversionResultData> {
    const {
      taskId,
      filename,
      filePath,
      outputDir = getDefaultResultsDir(),
      extractMedia = true,
      includeAnimations = true,
      includeNotes = true,
      targetVersion = 'latest',
      skipValidation = false,
      ignoreEncryption = false,
      onProgress,
    } = options;

    logger.info('Starting PPTX conversion', {
      taskId,
      filename,
      filePath,
      skipValidation,
      ignoreEncryption,
    });

    const startTime = new Date();

    try {
      // Step 1: Validate PPTX file
      let validationResult: any = { warnings: [] };
      if (!skipValidation) {
        onProgress?.(5, 'Validating PPTX file...');
        validationResult = await validatePPTXFile(filePath, {
          ignoreEncryption,
        });

        if (!validationResult.valid) {
          throw new PPTXValidationError(
            `Invalid PPTX file: ${validationResult.errors.map((e: any) => e.message).join(', ')}`,
            'VALIDATION_FAILED'
          );
        }
      } else {
        logger.info('Skipping PPTX validation');
      }

      // Step 2: Unzip PPTX file
      onProgress?.(15, 'Extracting PPTX content...');
      const extracted = await unzipPPTX(filePath, { ignoreEncryption });

      logger.info('PPTX extracted successfully', {
        slideCount: extracted.slides.size,
        mediaCount: extracted.media.size,
      });

      // Step 3: Extract and process media files
      let mediaFiles = new Map();
      if (extractMedia) {
        onProgress?.(30, 'Processing media files...');
        mediaFiles = await extractMediaFiles(extracted, {
          taskId,
          outputDir: getDefaultMediaDir(),
        });
      }

      // Step 4: Build relationship map for media resolution
      const relationshipMap = this.buildRelationshipMap(extracted);

      // Step 5: Create conversion context with relationship support
      const context = this.createConversionContext(mediaFiles, relationshipMap, targetVersion);

      // Step 6: Convert slides
      onProgress?.(50, 'Converting slides...');
      const orchestrator = new ConversionOrchestrator({
        includeAnimations,
        includeNotes,
        preserveZIndex: true,
        processGroups: true,
        targetVersion,
      });

      const presentation = await orchestrator.convert(extracted, context);

      // Step 7: Collect metadata
      onProgress?.(80, 'Collecting metadata...');
      const metadata = collectMetadata(presentation, extracted, filename, startTime);

      // Step 8: Collect warnings
      const warnings = collectWarnings(extracted, validationResult.warnings as any);

      // Step 9: Serialize result
      onProgress?.(90, 'Serializing result...');
      const jsonResult = serializeResult(presentation, metadata, warnings);

      // Step 10: Save result to file
      await this.saveResult(jsonResult, taskId, outputDir);

      onProgress?.(100, 'Conversion completed successfully');

      logger.info('PPTX conversion completed successfully', {
        taskId,
        slideCount: presentation.slides.length,
        elementCount: presentation.slides.reduce((sum, s) => sum + s.elements.length, 0),
        warningCount: warnings.length,
      });

      return {
        success: true,
        data: jsonResult,
        metadata,
        warnings,
      };
    } catch (error) {
      logger.error('PPTX conversion failed', {
        taskId,
        filename,
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });

      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error occurred',
      };
    }
  }

  /**
   * Build relationship map from extracted PPTX relationships
   * Maps slideIndex -> rId -> target filename
   */
  private buildRelationshipMap(extracted: ExtractedPPTX): Map<number, Map<string, string>> {
    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '',
    });

    const slideRelationships = new Map<number, Map<string, string>>();

    for (const [slideIndex, relXml] of extracted.relationships.entries()) {
      const relMap = new Map<string, string>();

      try {
        const parsed = parser.parse(relXml);
        const relationships = parsed['Relationships']?.['Relationship'] || [];
        const rels = Array.isArray(relationships) ? relationships : [relationships];

        for (const rel of rels) {
          if (rel && rel['Id'] && rel['Target']) {
            // Extract filename from target path (e.g., "../media/image1.png" -> "image1.png")
            const target = rel['Target'];
            const filename = path.basename(target);
            relMap.set(rel['Id'], filename);
          }
        }

        slideRelationships.set(slideIndex, relMap);
      } catch (error) {
        logger.warn(`Failed to parse relationships for slide ${slideIndex}`, {
          error: error instanceof Error ? error.message : String(error),
        });
      }
    }

    logger.debug('Built relationship map', {
      slideCount: slideRelationships.size,
    });

    return slideRelationships;
  }

  /**
   * Create conversion context with relationship-based media resolution
   */
  private createConversionContext(
    mediaFiles: Map<string, any>,
    relationshipMap: Map<number, Map<string, string>>,
    targetVersion: string
  ): ConversionContext {
    return {
      version: targetVersion,
      basePath: '',
      mediaFiles: new Map(),
      // Track current slide index for media resolution
      _currentSlideIndex: 1,
      _relationshipMap: relationshipMap,
      resolveMediaReference: (ref: string) => {
        // ref is a relationship ID like "rId1"
        // First, try to resolve using the relationship map
        const currentSlideIndex = (this as any)._currentSlideIndex || 1;
        const slideRelMap = relationshipMap.get(currentSlideIndex);

        if (slideRelMap) {
          const filename = slideRelMap.get(ref);
          if (filename && mediaFiles.has(filename)) {
            return mediaFiles.get(filename);
          }
        }

        // Fallback: try direct lookup by ref (in case ref is already a filename)
        if (mediaFiles.has(ref)) {
          return mediaFiles.get(ref);
        }

        // Second fallback: try all slides' relationships
        for (const [, relMap] of relationshipMap.entries()) {
          const filename = relMap.get(ref);
          if (filename && mediaFiles.has(filename)) {
            return mediaFiles.get(filename);
          }
        }

        logger.warn('Failed to resolve media reference', { ref });
        return null;
      },
      slideSize: {
        width: 1280,
        height: 720,
      },
    } as ConversionContext;
  }

  /**
   * Save result to file
   */
  private async saveResult(
    result: any,
    taskId: string,
    outputDir: string
  ): Promise<void> {
    try {
      await fs.mkdir(outputDir, { recursive: true });

      const jsonPath = `${outputDir}/${taskId}.json`;
      await fs.writeFile(jsonPath, JSON.stringify(result, null, 2), 'utf-8');

      logger.debug('Conversion result saved', {
        taskId,
        path: jsonPath,
      });
    } catch (error) {
      logger.error('Failed to save conversion result', {
        taskId,
        error: error instanceof Error ? error.message : String(error),
      });
      // Don't throw - result is still in memory
    }
  }
}

/**
 * Create conversion service instance and run conversion
 *
 * @param options - Conversion options
 * @returns Conversion result
 */
export async function runConversion(
  options: ConversionServiceOptions
): Promise<ConversionResultData> {
  const service = new ConversionService();
  return service.convert(options);
}
