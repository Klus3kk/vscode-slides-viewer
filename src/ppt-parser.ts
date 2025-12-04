import * as CFB from "cfb";

// PowerPoint record types (why am i doing this to myself...)
const RecordType = {
  RT_Document: 0x03E8,
  RT_Slide: 0x03EE,
  RT_SlideAtom: 0x03EF,
  RT_Notes: 0x03F0,
  RT_NotesAtom: 0x03F1,
  RT_MainMaster: 0x03F8,
  RT_SlideShowSlideInfoAtom: 0x03F9,
  RT_TextCharsAtom: 0x0FA0,
  RT_TextBytesAtom: 0x0FA8,
  RT_TextHeaderAtom: 0x0F9F,
  RT_TextSpecInfoAtom: 0x0FAA,
  RT_StyleTextPropAtom: 0x0FA1,
  RT_MasterTextPropAtom: 0x0FA2,
  RT_TextRulerAtom: 0x0FA3,
  RT_TextBookmarkAtom: 0x0FA7,
  RT_Drawing: 0x040C,
  RT_ProgTags: 0x1388,
  RT_ShapeAtom: 0x0F0A,
  RT_Shape: 0x0F00,
  RT_SlideListWithText: 0x0FF0,
  RT_TextBytesAtom_1: 0x0FA8,
  RT_ColorSchemeAtom: 0x07F0,
  RT_SlideSchemeColorSchemeAtom: 0x0FF1,
  RT_ExObjList: 0x0409,
  RT_ExObjListAtom: 0x040A,
  RT_PPDrawingGroup: 0x040B,
  RT_PPDrawing: 0x040C,
  RT_OfficeArtDggContainer: 0xF000,
  RT_OfficeArtSpgrContainer: 0xF003,
  RT_OfficeArtSpContainer: 0xF004,
  RT_OfficeArtClientTextbox: 0xF00D,
  RT_OfficeArtClientData: 0xF011,
};

interface RecordHeader {
  recVer: number;
  recInstance: number;
  recType: number;
  recLen: number;
}

interface TextRun {
  text: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  fontFamily?: string;
}

interface PptShape {
  type: string;
  box: { x: number; y: number; cx: number; cy: number };
  fill?: { type: string; color?: string };
  textData?: {
    paragraphs: Array<{
      align: string;
      runs: TextRun[];
      bullet?: any;
      level: number;
      marL: number;
      indent: number;
    }>;
    verticalAlign: string;
  };
  geom?: string | null;
  isMaster: boolean;
}

interface PptSlide {
  path: string;
  size: { cx: number; cy: number };
  shapes: PptShape[];
  background?: { color: string };
}

export async function parsePptFile(buffer: Buffer): Promise<PptSlide[]> {
  try {
    // Parse the Compound File Binary format
    const cfb = CFB.read(buffer, { type: "buffer" });
    
    // Find the PowerPoint Document stream
    const pptStream = findStream(cfb, "PowerPoint Document");
    if (!pptStream) {
      throw new Error("Not a valid PowerPoint file - PowerPoint Document stream not found");
    }

    // Find Current User stream for slide list
    const currentUserStream = findStream(cfb, "Current User");
    
    // Parse the document
    const parser = new PptParser(pptStream, currentUserStream);
    return parser.parse();
  } catch (error) {
    console.error("Error parsing .ppt file:", error);
    throw error;
  }
}

function findStream(cfb: CFB.CFB$Container, name: string): Uint8Array | null {
  for (const entry of cfb.FileIndex) {
    if (entry.name === name && entry.content) {
      const content = entry.content as unknown;
      if (content instanceof Uint8Array) return content;
      if (Array.isArray(content)) return Uint8Array.from(content);
      return null;
    }
  }
  return null;
}

class PptParser {
  private stream: Uint8Array;
  private currentUserStream: Uint8Array | null;
  private offset: number = 0;
  private slides: PptSlide[] = [];
  private slideWidth: number = 9144000; // Default: 10 inches in EMUs
  private slideHeight: number = 6858000; // Default: 7.5 inches in EMUs
  private textBySlide: Map<number, string[]> = new Map();
  private shapesBySlide: Map<number, PptShape[]> = new Map();

  constructor(stream: Uint8Array, currentUserStream: Uint8Array | null) {
    this.stream = stream;
    this.currentUserStream = currentUserStream;
  }

  parse(): PptSlide[] {
    this.offset = 0;
    
    // First pass: collect document structure
    this.parseDocument();
    
    // Build slides from collected data
    return this.buildSlides();
  }

  private parseDocument() {
    while (this.offset < this.stream.length - 8) {
      const header = this.readRecordHeader();
      const recordStart = this.offset;
      const recordEnd = this.offset + header.recLen;

      switch (header.recType) {
        case RecordType.RT_Document:
          this.parseDocumentContainer(recordEnd);
          break;
        
        case RecordType.RT_Slide:
          this.parseSlideContainer(recordEnd);
          break;
        
        case RecordType.RT_SlideAtom:
          this.parseSlideAtom();
          break;
        
        case RecordType.RT_MainMaster:
          this.parseMasterContainer(recordEnd);
          break;
        
        case RecordType.RT_SlideListWithText:
          this.parseSlideListWithText(recordEnd);
          break;

        default:
          // Skip unknown records
          this.offset = recordEnd;
          break;
      }

      // Safety: ensure we don't get stuck
      if (this.offset < recordEnd) {
        this.offset = recordEnd;
      }
    }
  }

  private parseDocumentContainer(endOffset: number) {
    while (this.offset < endOffset - 8) {
      const header = this.readRecordHeader();
      const recordEnd = this.offset + header.recLen;

      // Look for slide size information
      if (header.recType === 0x03F2) { // RT_Environment
        this.parseEnvironment(recordEnd);
      } else if (header.recType === RecordType.RT_SlideListWithText) {
        this.parseSlideListWithText(recordEnd);
      }

      this.offset = recordEnd;
    }
  }

  private parseEnvironment(endOffset: number) {
    // Parse slide size from environment
    while (this.offset < endOffset - 8) {
      const header = this.readRecordHeader();
      const recordEnd = this.offset + header.recLen;
      
      // Look for slide size atom (0x03F4)
      if (header.recType === 0x03F4 && header.recLen >= 8) {
        const view = new DataView(this.stream.buffer, this.offset, 8);
        this.slideWidth = view.getInt32(0, true);
        this.slideHeight = view.getInt32(4, true);
      }

      this.offset = recordEnd;
    }
  }

  private parseSlideContainer(endOffset: number) {
    const slideIndex = this.slides.length;
    const shapes: PptShape[] = [];

    while (this.offset < endOffset - 8) {
      const header = this.readRecordHeader();
      const recordEnd = this.offset + header.recLen;

      if (header.recType === RecordType.RT_PPDrawing) {
        const slideShapes = this.parseDrawingContainer(recordEnd);
        shapes.push(...slideShapes);
      } else if (header.recType === RecordType.RT_ColorSchemeAtom) {
        // Parse color scheme if needed
      }

      this.offset = recordEnd;
    }

    this.shapesBySlide.set(slideIndex, shapes);
    this.slides.push({
      path: `slide${slideIndex}`,
      size: { cx: this.slideWidth, cy: this.slideHeight },
      shapes: shapes,
    });
  }

  private parseSlideAtom() {
    // SlideAtom contains slide properties
    // Skip for now as we're handling basics
    const view = new DataView(this.stream.buffer, this.offset, 24);
    // Contains geometry, flags, etc.
  }

  private parseMasterContainer(endOffset: number) {
    // Parse master slide (similar to regular slide)
    while (this.offset < endOffset - 8) {
      const header = this.readRecordHeader();
      this.offset += header.recLen;
    }
  }

  private parseSlideListWithText(endOffset: number) {
    const slideIndex = this.textBySlide.size;
    const texts: string[] = [];

    while (this.offset < endOffset - 8) {
      const header = this.readRecordHeader();
      const recordEnd = this.offset + header.recLen;

      if (header.recType === RecordType.RT_TextCharsAtom) {
        const text = this.readTextCharsAtom(header.recLen);
        texts.push(text);
      } else if (header.recType === RecordType.RT_TextBytesAtom) {
        const text = this.readTextBytesAtom(header.recLen);
        texts.push(text);
      }

      this.offset = recordEnd;
    }

    if (texts.length > 0) {
      this.textBySlide.set(slideIndex, texts);
    }
  }

  private parseDrawingContainer(endOffset: number): PptShape[] {
    const shapes: PptShape[] = [];

    while (this.offset < endOffset - 8) {
      const header = this.readRecordHeader();
      const recordEnd = this.offset + header.recLen;

      // OfficeArt containers
      if (header.recType === 0xF003 || header.recType === 0xF004) {
        const containerShapes = this.parseShapeContainer(recordEnd);
        shapes.push(...containerShapes);
      }

      this.offset = recordEnd;
    }

    return shapes;
  }

  private parseShapeContainer(endOffset: number): PptShape[] {
    const shapes: PptShape[] = [];
    let shapeData: any = {};

    while (this.offset < endOffset - 8) {
      const header = this.readRecordHeader();
      const recordEnd = this.offset + header.recLen;

      if (header.recType === 0xF00A) {
        // OfficeArtFSP - shape properties
        shapeData = this.parseShapeProperties();
      } else if (header.recType === 0xF00B) {
        // OfficeArtFOPT - shape formatting
        const props = this.parseShapeFormatting(header.recLen);
        shapeData = { ...shapeData, ...props };
      } else if (header.recType === 0xF00D) {
        // OfficeArtClientTextbox - contains text
        const text = this.parseClientTextbox(recordEnd);
        if (text) shapeData.text = text;
      } else if (header.recType === 0xF004) {
        // Nested shape container
        const nestedShapes = this.parseShapeContainer(recordEnd);
        shapes.push(...nestedShapes);
      }

      this.offset = recordEnd;
    }

    // Convert collected data into shape
    if (shapeData.bounds) {
      const shape = this.createShape(shapeData);
      if (shape) shapes.push(shape);
    }

    return shapes;
  }

  private parseShapeProperties(): any {
    if (this.offset + 8 > this.stream.length) return {};
    
    const view = new DataView(this.stream.buffer, this.offset, 8);
    const shapeId = view.getUint32(0, true);
    const flags = view.getUint32(4, true);

    return { shapeId, flags };
  }

  private parseShapeFormatting(length: number): any {
    const props: any = { bounds: null, fill: null };
    const numProperties = Math.floor(length / 6);

    for (let i = 0; i < numProperties; i++) {
      const propOffset = this.offset + i * 6;
      if (propOffset + 6 > this.stream.length) break;

      const view = new DataView(this.stream.buffer, propOffset, 6);
      const propId = view.getUint16(0, true);
      const propValue = view.getUint32(2, true);

      // Shape bounds (0x0004 = left, 0x0005 = top, 0x0006 = right, 0x0007 = bottom)
      if (propId === 0x0004) {
        if (!props.bounds) props.bounds = {};
        props.bounds.left = propValue;
      } else if (propId === 0x0005) {
        if (!props.bounds) props.bounds = {};
        props.bounds.top = propValue;
      } else if (propId === 0x0006) {
        if (!props.bounds) props.bounds = {};
        props.bounds.right = propValue;
      } else if (propId === 0x0007) {
        if (!props.bounds) props.bounds = {};
        props.bounds.bottom = propValue;
      }
      // Fill color (0x0181)
      else if (propId === 0x0181) {
        props.fill = { color: this.colorFromInt(propValue) };
      }
    }

    return props;
  }

  private parseClientTextbox(endOffset: number): string | null {
    const texts: string[] = [];

    while (this.offset < endOffset - 8) {
      const header = this.readRecordHeader();
      const recordEnd = this.offset + header.recLen;

      if (header.recType === RecordType.RT_TextCharsAtom) {
        texts.push(this.readTextCharsAtom(header.recLen));
      } else if (header.recType === RecordType.RT_TextBytesAtom) {
        texts.push(this.readTextBytesAtom(header.recLen));
      }

      this.offset = recordEnd;
    }

    return texts.length > 0 ? texts.join("\n") : null;
  }

  private createShape(data: any): PptShape | null {
    if (!data.bounds) return null;

    const { left = 0, top = 0, right = 0, bottom = 0 } = data.bounds;
    
    // Convert from EMUs to pixels (approximation)
    const emusToPixels = (emus: number) => Math.round(emus / 9525);

    const shape: PptShape = {
      type: data.text ? "text" : "shape",
      box: {
        x: emusToPixels(left),
        y: emusToPixels(top),
        cx: emusToPixels(right - left),
        cy: emusToPixels(bottom - top),
      },
      isMaster: false,
    };

    if (data.fill) {
      shape.fill = { type: "solid", color: data.fill.color };
    }

    if (data.text) {
      shape.textData = {
        paragraphs: [{
          align: "left",
          runs: [{ text: data.text }],
          level: 0,
          marL: 0,
          indent: 0,
        }],
        verticalAlign: "center",
      };
    }

    return shape;
  }

  private readRecordHeader(): RecordHeader {
    if (this.offset + 8 > this.stream.length) {
      throw new Error("Unexpected end of stream while reading record header");
    }

    const view = new DataView(this.stream.buffer, this.offset, 8);
    
    const verAndInstance = view.getUint16(0, true);
    const recVer = verAndInstance & 0x0F;
    const recInstance = (verAndInstance >> 4) & 0x0FFF;
    const recType = view.getUint16(2, true);
    const recLen = view.getUint32(4, true);

    this.offset += 8;

    return { recVer, recInstance, recType, recLen };
  }

  private readTextCharsAtom(length: number): string {
    if (this.offset + length > this.stream.length) {
      return "";
    }

    const chars: string[] = [];
    for (let i = 0; i < length; i += 2) {
      const charCode = this.stream[this.offset + i] | (this.stream[this.offset + i + 1] << 8);
      chars.push(String.fromCharCode(charCode));
    }

    return chars.join("");
  }

  private readTextBytesAtom(length: number): string {
    if (this.offset + length > this.stream.length) {
      return "";
    }

    const decoder = new TextDecoder("windows-1252");
    const bytes = this.stream.slice(this.offset, this.offset + length);
    return decoder.decode(bytes);
  }

  private colorFromInt(colorInt: number): string {
    const r = colorInt & 0xFF;
    const g = (colorInt >> 8) & 0xFF;
    const b = (colorInt >> 16) & 0xFF;
    return `#${r.toString(16).padStart(2, "0")}${g.toString(16).padStart(2, "0")}${b.toString(16).padStart(2, "0")}`;
  }

  private buildSlides(): PptSlide[] {
    // Merge text and shape data
    this.slides.forEach((slide, index) => {
      const texts = this.textBySlide.get(index);
      if (texts && texts.length > 0) {
        // Add text to first shape or create new text shape
        if (slide.shapes.length === 0) {
          slide.shapes.push({
            type: "text",
            box: { x: 50, y: 50, cx: 800, cy: 400 },
            textData: {
              paragraphs: texts.map(text => ({
                align: "left",
                runs: [{ text }],
                level: 0,
                marL: 0,
                indent: 0,
              })),
              verticalAlign: "flex-start",
            },
            isMaster: false,
          });
        }
      }
    });

    return this.slides;
  }
}
