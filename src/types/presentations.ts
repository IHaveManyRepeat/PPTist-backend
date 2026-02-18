/**
 * Presentation Type Definitions
 *
 * Defines types for presentation data structures.
 *
 * @module types/presentations
 */

/**
 * Base element interface
 */
export interface BaseElement {
  id: string;
  type: string;
  name?: string;
  hidden?: boolean;
  locked?: boolean;
  zIndex?: number;
}

/**
 * Slide interface
 */
export interface Slide {
  id: string;
  elements: SlideElement[];
  background?: SlideBackground;
  notes?: string;
  layoutId?: string;
  index?: number;
}

/**
 * Slide element union type
 */
export type SlideElement =
  | TextSlideElement
  | ImageSlideElement
  | ShapeSlideElement
  | LineSlideElement
  | ChartSlideElement
  | TableSlideElement
  | VideoSlideElement
  | AudioSlideElement
  | GroupSlideElement;

/**
 * Text slide element
 */
export interface TextSlideElement extends BaseElement {
  type: 'text';
  x: number;
  y: number;
  width: number;
  height: number;
  rotate?: number;
  text?: string;
  content?: string;
  defaultValue?: string;
  fill?: any;
  outline?: any;
}

/**
 * Image slide element
 */
export interface ImageSlideElement extends BaseElement {
  type: 'image';
  x: number;
  y: number;
  width: number;
  height: number;
  rotate?: number;
  src: string;
  crop?: any;
}

/**
 * Shape slide element
 */
export interface ShapeSlideElement extends BaseElement {
  type: 'shape';
  x: number;
  y: number;
  width: number;
  height: number;
  rotate?: number;
  fill?: any;
  outline?: any;
  text?: string;
  shape?: any;
}

/**
 * Line slide element
 */
export interface LineSlideElement extends BaseElement {
  type: 'line';
  x: number;
  y: number;
  width: number;
  height: number;
  startX: number;
  startY: number;
  endX: number;
  endY: number;
  color?: string;
  style?: string;
}

/**
 * Chart slide element
 */
export interface ChartSlideElement extends BaseElement {
  type: 'chart';
  x: number;
  y: number;
  width: number;
  height: number;
  rotate?: number;
  data?: any;
  options?: any;
}

/**
 * Table slide element
 */
export interface TableSlideElement extends BaseElement {
  type: 'table';
  x: number;
  y: number;
  width: number;
  height: number;
  rotate?: number;
  rowCount?: number;
  colCount?: number;
  data?: any;
}

/**
 * Video slide element
 */
export interface VideoSlideElement extends BaseElement {
  type: 'video';
  x: number;
  y: number;
  width: number;
  height: number;
  src: string;
  autoplay?: boolean;
  loop?: boolean;
  controls?: boolean;
}

/**
 * Audio slide element
 */
export interface AudioSlideElement extends BaseElement {
  type: 'audio';
  x: number;
  y: number;
  width: number;
  height: number;
  src: string;
  autoplay?: boolean;
  loop?: boolean;
  controls?: boolean;
}

/**
 * Group slide element
 */
export interface GroupSlideElement extends BaseElement {
  type: 'group';
  x: number;
  y: number;
  width: number;
  height: number;
  elements: SlideElement[];
}

/**
 * Slide background
 */
export interface SlideBackground {
  type?: 'solid' | 'gradient' | 'image' | 'pattern';
  color?: string;
  gradientColors?: string[];
  imageRef?: string;
  value?: any;
}

/**
 * Presentation interface
 */
export interface Presentation {
  id?: string;
  title?: string;
  author?: string;
  width: number;
  height: number;
  slides: Slide[];
  metadata?: Record<string, unknown>;
}

/**
 * Theme interface
 */
export interface Theme {
  width: number;
  height: number;
  fontName: string;
  fontColor: string;
}
