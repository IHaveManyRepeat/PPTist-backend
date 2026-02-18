/**
 * Conversion Task Worker
 *
 * Worker that processes conversion tasks from the queue.
 * Handles progress updates, error recovery, and result storage.
 *
 * @module services/conversion/worker
 */

import type { Task, TaskQueue } from '../../types/queue';
import { TaskStatus } from '../../types/queue';
import { runConversion } from './index.js';
import { createZIPArchive } from '../storage/zip-creator';
import { logger } from '../../utils/logger';
import { promises as fs } from 'fs';
import * as path from 'path';
import {
  recordConversionStart,
  recordConversionComplete,
  updateQueueMetrics,
} from '../../utils/metrics-enhanced';
import { getConversionLogger, ConversionStep } from '../../utils/conversion-logger';
import { getMemoryMonitor, formatMemoryStats } from '../../utils/memory-monitor';
import { getDefaultResultsDir, getDefaultMediaDir } from '../../utils/paths';

/**
 * Worker options
 */
export interface ConversionWorkerOptions {
  /**
   * Task queue instance
   */
  queue: TaskQueue;

  /**
   * Results directory
   */
  resultsDir?: string;

  /**
   * Media directory
   */
  mediaDir?: string;

  /**
   * Worker ID for logging
   */
  workerId?: string;
}

/**
 * Conversion worker class
 */
export class ConversionWorker {
  private queue: TaskQueue;
  private resultsDir: string;
  private mediaDir: string;
  private workerId: string;
  private isRunning: boolean = false;
  private conversionLogger = getConversionLogger();
  private memoryMonitor = getMemoryMonitor();

  constructor(options: ConversionWorkerOptions) {
    this.queue = options.queue;
    this.resultsDir = options.resultsDir || getDefaultResultsDir();
    this.mediaDir = options.mediaDir || getDefaultMediaDir();
    this.workerId = options.workerId || 'worker-1';

    // Ensure directories exist
    this.initializeDirectories();
  }

  /**
   * Initialize directories
   */
  private async initializeDirectories(): Promise<void> {
    try {
      await fs.mkdir(this.resultsDir, { recursive: true });
      await fs.mkdir(this.mediaDir, { recursive: true });
    } catch (error) {
      logger.error('Failed to initialize worker directories', {
        error: error instanceof Error ? error.message : String(error),
      });
    }
  }

  /**
   * Start worker
   */
  async start(): Promise<void> {
    if (this.isRunning) {
      logger.warn('Worker is already running', { workerId: this.workerId });
      return;
    }

    this.isRunning = true;
    logger.info('Conversion worker started', { workerId: this.workerId });

    // Register task handler with queue
    await this.queue.start(this.taskHandler.bind(this));

    // Start periodic queue metrics updates
    this.startQueueMetricsUpdater();

    logger.info('Worker task handler registered', { workerId: this.workerId });
  }

  /**
   * Stop worker
   */
  async stop(): Promise<void> {
    this.isRunning = false;
    logger.info('Conversion worker stopped', { workerId: this.workerId });
  }

  /**
   * Start periodic queue metrics updates
   */
  private startQueueMetricsUpdater(): void {
    // Update metrics every 10 seconds
    const intervalId = setInterval(() => {
      if (this.isRunning) {
        updateQueueMetrics(this.queue);
      }
    }, 10000);

    // Clear interval on process exit
    process.on('beforeExit', () => clearInterval(intervalId));
    process.on('SIGINT', () => clearInterval(intervalId));
    process.on('SIGTERM', () => clearInterval(intervalId));
  }

  /**
   * Process tasks from queue
   */
  private async processTasks(): Promise<void> {
    while (this.isRunning) {
      try {
        // Get next task from queue (blocking)
        const task = await this.getNextTask();

        if (task) {
          await this.processTask(task);
        } else {
          // No task available, wait before retrying
          await this.sleep(1000);
        }
      } catch (error) {
        logger.error('Error in task processing loop', {
          workerId: this.workerId,
          error: error instanceof Error ? error.message : String(error),
        });

        // Wait before retrying
        await this.sleep(5000);
      }
    }
  }

  /**
   * Get next task from queue
   */
  private async getNextTask(): Promise<Task | null> {
    try {
      // Check if we can accept new conversion
      const canAccept = this.memoryMonitor.canAcceptConversion();
      if (!canAccept.allowed) {
        logger.debug('Cannot accept new conversion', {
          reason: canAccept.reason,
          memoryStats: formatMemoryStats(this.memoryMonitor.getStats()),
        });
        return null;
      }

      // Get a task in 'queued' status
      const tasks = await this.queue.getTasksByStatus(TaskStatus.QUEUED);

      if (tasks.length === 0) {
        return null;
      }

      // Get the first task
      const task = tasks[0];

      // Update task status to processing
      await this.queue.updateTask(task.id, {
        status: TaskStatus.PROCESSING,
        progress: 0,
      });

      return task;
    } catch (error) {
      logger.error('Failed to get next task', {
        workerId: this.workerId,
        error: error instanceof Error ? error.message : String(error),
      });
      return null;
    }
  }

  /**
   * Task handler for queue
   * Called by the queue when a task is ready to be processed
   */
  private async taskHandler(taskData: any, taskId: string): Promise<any> {
    logger.info('Task handler called', {
      workerId: this.workerId,
      taskId,
    });

    // Create task object
    const task: Task = {
      id: taskId,
      data: taskData,
      status: TaskStatus.PROCESSING,
      progress: 0,
      result: null,
      error: null,
      metadata: {
        createdAt: new Date(),
        startedAt: new Date(),
        retryCount: 0,
        priority: 'normal' as any,
      },
    };

    // Process the task
    const success = await this.processTask(task);

    // If task failed, throw an error so the queue correctly marks it as FAILED
    // This prevents the race condition where queue overwrites FAILED status with COMPLETED
    if (!success) {
      throw new Error(task.error || 'Conversion failed');
    }

    // Return the result
    return task.result;
  }

  /**
   * Process single task
   * @returns true if task completed successfully, false otherwise
   */
  private async processTask(task: Task): Promise<boolean> {
    logger.info('Processing conversion task', {
      workerId: this.workerId,
      taskId: task.id,
      filename: (task.data as any).originalFilename,
    });

    // Register conversion with memory monitor
    this.memoryMonitor.registerConversionStart();

    // Start conversion logging
    await this.conversionLogger.startConversion(task.id, {
      filename: (task.data as any).originalFilename,
      size: (task.data as any).metadata?.size || 0,
    });

    // Record conversion start
    recordConversionStart(task.id);

    try {
      // Start validation step
      this.conversionLogger.startStep(task.id, ConversionStep.VALIDATION);

      // Define progress callback
      const onProgress = (progress: number, message: string) => {
        this.updateTaskProgress(task.id, progress, message);
        this.conversionLogger.log({
          taskId: task.id,
          step: ConversionStep.CONVERSION,
          message,
          progress,
        });
      };

      // End validation step
      await this.conversionLogger.endStep(
        task.id,
        ConversionStep.VALIDATION,
        'Validation complete'
      );

      // Start conversion step
      this.conversionLogger.startStep(task.id, ConversionStep.CONVERSION);

      // Run conversion
      const taskOptions = (task.data as any).options || {};
      const result = await runConversion({
        taskId: task.id,
        filename: (task.data as any).originalFilename,
        filePath: (task.data as any).filePath,
        outputDir: this.resultsDir,
        extractMedia: true,
        includeAnimations: true,
        includeNotes: true,
        targetVersion: 'latest',
        ignoreEncryption: taskOptions.ignoreEncryption === true,
        onProgress,
      });

      // End conversion step
      await this.conversionLogger.endStep(
        task.id,
        ConversionStep.CONVERSION,
        'Conversion complete',
        { slideCount: result.data?.slides?.length }
      );

      // Record conversion completion
      recordConversionComplete(
        task.id,
        result.success ? 'success' : 'error',
        (task.data as any).metadata?.size
      );

      if (result.success && result.data) {
        // Task completed successfully
        const taskUpdate: any = {
          status: TaskStatus.COMPLETED,
          progress: 100,
          result: result.data,
          warnings: result.warnings?.map(w => `${w.code}: ${w.message}`),
        };

        await this.queue.updateTask(task.id, taskUpdate);

        // Update local task result for return value
        task.result = result.data;

        // Log warnings
        if (result.warnings && result.warnings.length > 0) {
          for (const warning of result.warnings) {
            await this.conversionLogger.addWarning(
              task.id,
              `${warning.code}: ${warning.message}`
            );
          }
        }

        // Create ZIP archive
        this.conversionLogger.startStep(task.id, ConversionStep.SERIALIZATION);
        await this.createTaskZIP(task.id, (task.data as any).originalFilename);
        await this.conversionLogger.endStep(
          task.id,
          ConversionStep.SERIALIZATION,
          'ZIP archive created'
        );

        // Complete conversion logging
        await this.conversionLogger.completeConversion(task.id, true);

        logger.info('Task completed successfully', {
          workerId: this.workerId,
          taskId: task.id,
        });

        return true;
      } else {
        // Task failed
        await this.conversionLogger.addError(
          task.id,
          result.error || 'Unknown error'
        );

        await this.queue.updateTask(task.id, {
          status: TaskStatus.FAILED,
          progress: 0,
          error: result.error || 'Unknown error',
        });

        // Complete conversion logging
        await this.conversionLogger.completeConversion(task.id, false);

        logger.error('Task failed', {
          workerId: this.workerId,
          taskId: task.id,
          error: result.error,
        });

        // Set error for handler to throw
        task.error = result.error || 'Unknown error';
        return false;
      }

      // Clean up uploaded file
      await this.cleanupUploadedFile((task.data as any).filePath);
    } catch (error) {
      // Task failed with exception
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';

      await this.conversionLogger.addError(task.id, errorMessage);

      await this.queue.updateTask(task.id, {
        status: TaskStatus.FAILED,
        progress: 0,
        error: errorMessage,
      });

      // Complete conversion logging
      await this.conversionLogger.completeConversion(task.id, false);

      logger.error('Task processing failed', {
        workerId: this.workerId,
        taskId: task.id,
        error: errorMessage,
        stack: error instanceof Error ? error.stack : undefined,
      });

      // Set error for handler to throw
      task.error = errorMessage;

      // Clean up uploaded file
      await this.cleanupUploadedFile((task.data as any).filePath);

      return false;
    } finally {
      // Unregister conversion with memory monitor
      this.memoryMonitor.registerConversionComplete();

      // Clear log summary from memory after a delay
      setTimeout(() => {
        this.conversionLogger.clearSummary(task.id);
      }, 60000); // Keep for 1 minute
    }
  }

  /**
   * Update task progress
   */
  private async updateTaskProgress(
    taskId: string,
    progress: number,
    message: string
  ): Promise<void> {
    try {
      await this.queue.updateTask(taskId, {
        progress: Math.round(progress),
      });

      logger.debug('Task progress updated', {
        taskId,
        progress,
        message,
      });
    } catch (error) {
      logger.error('Failed to update task progress', {
        taskId,
        error: error instanceof Error ? error.message : String(error),
      });
    }
  }

  /**
   * Create ZIP archive for task
   */
  private async createTaskZIP(taskId: string, filename: string): Promise<void> {
    try {
      const jsonPath = path.join(this.resultsDir, `${taskId}.json`);
      const mediaPath = path.join(this.mediaDir, taskId);
      const zipPath = path.join(this.resultsDir, `${taskId}.zip`);

      // Check if media directory exists
      let hasMedia = false;
      try {
        const mediaFiles = await fs.readdir(mediaPath);
        hasMedia = mediaFiles.length > 0;
      } catch {
        // Media directory doesn't exist
      }

      // Create ZIP archive
      await createZIPArchive({
        jsonPath,
        mediaPath: hasMedia ? mediaPath : undefined,
        outputPath: zipPath,
        filename: filename.replace('.pptx', '.json'),
      });

      logger.debug('Task ZIP archive created', {
        taskId,
        zipPath,
        hasMedia,
      });
    } catch (error) {
      logger.error('Failed to create task ZIP archive', {
        taskId,
        error: error instanceof Error ? error.message : String(error),
      });
      // Don't fail the task if ZIP creation fails
    }
  }

  /**
   * Clean up uploaded file
   */
  private async cleanupUploadedFile(filePath: string): Promise<void> {
    try {
      await fs.unlink(filePath);
      logger.debug('Uploaded file cleaned up', { filePath });
    } catch (error) {
      logger.warn('Failed to clean up uploaded file', {
        filePath,
        error: error instanceof Error ? error.message : String(error),
      });
    }
  }

  /**
   * Sleep utility
   */
  private sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
}

/**
 * Create and start conversion worker
 */
export async function createAndStartWorker(options: ConversionWorkerOptions): Promise<ConversionWorker> {
  const worker = new ConversionWorker(options);
  await worker.start();
  return worker;
}
