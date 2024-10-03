import { XMLParser, XMLBuilder } from "fast-xml-parser";
import JSZip from "jszip";

export class VisioToBpmnConverterClass {
  private bpmn_ns: string;
  private bpmndi_ns: string;
  private dc_ns: string;
  private di_ns: string;
  private xsi_ns: string;
  private collaborationId: string;
  private participantId: string;
  private processId: string;
  private bpmnJson: any; // Add this line

  constructor() {
    this.bpmn_ns = "http://www.omg.org/spec/BPMN/20100524/MODEL";
    this.bpmndi_ns = "http://www.omg.org/spec/BPMN/20100524/DI";
    this.dc_ns = "http://www.omg.org/spec/DD/20100524/DC";
    this.di_ns = "http://www.omg.org/spec/DD/20100524/DI";
    this.xsi_ns = "http://www.w3.org/2001/XMLSchema-instance";
    this.collaborationId = `Collaboration_${this._generateUniqueId()}`;
    this.participantId = `Participant_${this._generateUniqueId()}`;
    this.processId = `Process_${this._generateUniqueId()}`;
  }

  async convert(fileContent: ArrayBuffer): Promise<string> {
    try {
      const zip = new JSZip();
      await zip.loadAsync(fileContent);

      let xmlContent = await this._findAndExtractXml(zip);

      if (!xmlContent) {
        throw new Error("Could not find valid XML content in the Visio file");
      }

      xmlContent = this._preprocessXml(xmlContent);

      const parser = new XMLParser({
        ignoreAttributes: false,
        attributeNamePrefix: "@_",
        textNodeName: "#text",
        ignoreDeclaration: true,
        ignorePiTags: true,
      });

      const visioJson = parser.parse(xmlContent);
      console.log("Parsed Visio JSON:", JSON.stringify(visioJson, null, 2));

      this.bpmnJson = this._createBpmnStructure();

      this._convertShapes(visioJson);

      console.log("Final BPMN JSON structure:", JSON.stringify(this.bpmnJson, null, 2));

      const builder = new XMLBuilder({
        ignoreAttributes: false,
        format: true,
        suppressEmptyNode: true,
      });

      const xml = builder.build(this.bpmnJson);
      console.log("Generated XML:", xml);
      return xml;
    } catch (error) {
      console.error("Error in convert method:", error);
      throw new Error(`Failed to parse Visio file: ${error.message}`);
    }
  }

  private async _findAndExtractXml(zip: JSZip): Promise<string | null> {
    const possibleXmlFiles = ["visio/pages/page1.xml", "word/document.xml", "content.xml"];

    for (const fileName of possibleXmlFiles) {
      const file = zip.file(fileName);
      if (file) {
        return await file.async("string");
      }
    }

    // If no matching file is found, try to find any XML file
    const allFiles = Object.keys(zip.files);
    for (const fileName of allFiles) {
      if (fileName.endsWith(".xml")) {
        const file = zip.file(fileName);
        if (file) {
          return await file.async("string");
        }
      }
    }

    return null;
  }

  private _preprocessXml(xmlContent: string): string {
    // Remove any invalid characters or fix common issues
    xmlContent = xmlContent.replace(/&#x0;/g, ""); // Remove null characters
    xmlContent = xmlContent.replace(/&(?!(?:amp|lt|gt|quot|apos);)/g, "&amp;"); // Escape unescaped ampersands

    // Ensure the XML is well-formed
    if (!xmlContent.trim().startsWith("<?xml")) {
      xmlContent = '<?xml version="1.0" encoding="UTF-8"?>' + xmlContent;
    }

    return xmlContent;
  }

  private _createBpmnStructure() {
    const definitionsId = `Definitions_${this._generateUniqueId()}`;
    this.processId = `Process_${this._generateUniqueId()}`;
    this.participantId = `Participant_${this._generateUniqueId()}`;
    this.collaborationId = `Collaboration_${this._generateUniqueId()}`;

    return {
      "bpmn:definitions": {
        "@_xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
        "@_xmlns:bpmn": "http://www.omg.org/spec/BPMN/20100524/MODEL",
        "@_xmlns:bpmndi": "http://www.omg.org/spec/BPMN/20100524/DI",
        "@_xmlns:dc": "http://www.omg.org/spec/DD/20100524/DC",
        "@_xmlns:di": "http://www.omg.org/spec/DD/20100524/DI",
        "@_id": definitionsId,
        "@_targetNamespace": "http://bpmn.io/schema/bpmn",
        "@_exporter": "bpmn-js (https://demo.bpmn.io)",
        "@_exporterVersion": "9.0.3",
        "bpmn:collaboration": {
          "@_id": this.collaborationId,
        },
        "bpmn:process": {
          "@_id": this.processId,
          "@_isExecutable": "false",
        },
        "bpmndi:BPMNDiagram": {
          "@_id": `BPMNDiagram_${this._generateUniqueId()}`,
          "bpmndi:BPMNPlane": {
            "@_id": `BPMNPlane_${this._generateUniqueId()}`,
            "@_bpmnElement": this.collaborationId,
          },
        },
      },
    };
  }

  private _convertShapes(visioJson: any) {
    const shapes = this._findShapes(visioJson);
    console.log("Found shapes:", shapes);

    const scaleFactor = 100; // Increase this value to spread out the shapes more

    // Ensure the necessary structures exist
    if (!this.bpmnJson["bpmn:definitions"]["bpmn:process"]) {
      this.bpmnJson["bpmn:definitions"]["bpmn:process"] = {
        "@_id": this.processId,
        "@_isExecutable": "false",
      };
    }

    const process = this.bpmnJson["bpmn:definitions"]["bpmn:process"];
    const bpmndiPlane = this.bpmnJson["bpmn:definitions"]["bpmndi:BPMNDiagram"]["bpmndi:BPMNPlane"];

    // Ensure bpmndi:BPMNShape is initialized as an array
    if (!bpmndiPlane["bpmndi:BPMNShape"]) {
      bpmndiPlane["bpmndi:BPMNShape"] = [];
    } else if (!Array.isArray(bpmndiPlane["bpmndi:BPMNShape"])) {
      bpmndiPlane["bpmndi:BPMNShape"] = [bpmndiPlane["bpmndi:BPMNShape"]];
    }

    // First, find the pool/participant
    const poolShape = shapes.find(shape => this._determineShapeType(shape) === "bpmn:participant");
    if (!poolShape) {
      console.warn("No pool/participant found. Creating a default one.");
      this._createDefaultParticipant(process, bpmndiPlane, scaleFactor);
    } else {
      this._createParticipant(process, bpmndiPlane, poolShape, scaleFactor);
    }

    // Then, process all other shapes
    shapes.forEach((shape: any, index: number) => {
      console.log(`Processing shape ${index}:`, shape);
      const shapeType = this._determineShapeType(shape);
      console.log(`Determined shape type for shape ${index}:`, shapeType);
      if (shapeType && shapeType !== "bpmn:participant") {
        const bpmnElement = this._createBpmnElement(process, shapeType, shape);
        this._createBpmnDiElement(bpmndiPlane, bpmnElement, shape, scaleFactor);
      }
    });

    this._createSequenceFlows(shapes, process, bpmndiPlane, scaleFactor);
  }

  private _findShapes(visioJson: any): any[] {
    if (visioJson.PageContents && visioJson.PageContents.Shapes && visioJson.PageContents.Shapes.Shape) {
      return Array.isArray(visioJson.PageContents.Shapes.Shape) ? visioJson.PageContents.Shapes.Shape : [visioJson.PageContents.Shapes.Shape];
    }
    return [];
  }

  private _determineShapeType(shape: any): string | null {
    const text = this._getShapeText(shape).toLowerCase();

    if (text.includes("start")) return "bpmn:startEvent";
    if (text.includes("end")) return "bpmn:endEvent";
    if (text.includes("gateway")) return "bpmn:exclusiveGateway";
    if (text.includes("task") || text.includes("activity")) return "bpmn:task";
    if (text.includes("subprocess")) return "bpmn:subProcess";
    if (text.includes("lane")) return "bpmn:lane";
    if (text.includes("pool")) return "bpmn:participant";

    return "bpmn:task"; // Default to task if no specific type is identified
  }

  private _getShapeText(shape: any): string {
    if (shape.Text && typeof shape.Text === "string") {
      return shape.Text;
    }

    if (shape.Text && shape.Text.Cp && typeof shape.Text.Cp === "string") {
      return shape.Text.Cp;
    }

    if (shape.Text && Array.isArray(shape.Text.Cp)) {
      return shape.Text.Cp.join(" ");
    }

    if (shape.Text && shape.Text["#text"] && typeof shape.Text["#text"] === "string") {
      return shape.Text["#text"];
    }

    return "";
  }

  private _createBpmnElement(process: any, elementType: string, shape: any) {
    const elementId = `${elementType.split(":")[1]}_${this._generateUniqueId()}`;
    const elementName = this._getShapeText(shape);

    const element: any = { "@_id": elementId, "@_name": elementName };

    if (elementType === "bpmn:lane") {
      if (!process["bpmn:laneSet"]) {
        process["bpmn:laneSet"] = { "@_id": `LaneSet_${this._generateUniqueId()}`, "bpmn:lane": [] };
      }
      process["bpmn:laneSet"]["bpmn:lane"].push(element);
    } else {
      if (!process[elementType]) {
        process[elementType] = [];
      }
      process[elementType].push(element);
    }

    return element;
  }

  private _createBpmnDiElement(bpmndiPlane: any, bpmnElement: any, shape: any, scaleFactor: number) {
    const bounds = this._getShapeBounds(shape, scaleFactor);

    const bpmnShape = {
      "@_id": `BPMNShape_${this._generateUniqueId()}`,
      "@_bpmnElement": bpmnElement["@_id"],
      "dc:Bounds": bounds,
    };

    // Ensure bpmndi:BPMNShape is always an array
    if (!bpmndiPlane["bpmndi:BPMNShape"]) {
      bpmndiPlane["bpmndi:BPMNShape"] = [];
    } else if (!Array.isArray(bpmndiPlane["bpmndi:BPMNShape"])) {
      bpmndiPlane["bpmndi:BPMNShape"] = [bpmndiPlane["bpmndi:BPMNShape"]];
    }

    // Now we can safely push the new shape
    bpmndiPlane["bpmndi:BPMNShape"].push(bpmnShape);
  }

  private _getShapeBounds(shape: any, scaleFactor: number) {
    let x = 0,
      y = 0,
      width = 100,
      height = 80;

    if (Array.isArray(shape.Cell)) {
      shape.Cell.forEach((cell: any) => {
        switch (cell["@_N"]) {
          case "PinX":
            x = parseFloat(cell["@_V"]) || 0;
            break;
          case "PinY":
            y = parseFloat(cell["@_V"]) || 0;
            break;
          case "Width":
            width = parseFloat(cell["@_V"]) || 100;
            break;
          case "Height":
            height = parseFloat(cell["@_V"]) || 80;
            break;
        }
      });
    }

    return {
      "@_x": Math.round(x * scaleFactor),
      "@_y": Math.round(y * scaleFactor),
      "@_width": Math.round(width),
      "@_height": Math.round(height),
    };
  }

  private _createSequenceFlows(shapes: any[], process: any, bpmndiPlane: any, scaleFactor: number) {
    shapes.forEach((shape: any) => {
      if (shape.Connects && Array.isArray(shape.Connects.Connect)) {
        shape.Connects.Connect.forEach((connect: any) => {
          const sourceRef = connect["@_FromSheet"];
          const targetRef = connect["@_ToSheet"];

          const sequenceFlow = {
            "@_id": `SequenceFlow_${this._generateUniqueId()}`,
            "@_sourceRef": sourceRef,
            "@_targetRef": targetRef,
          };

          if (!process["bpmn:sequenceFlow"]) {
            process["bpmn:sequenceFlow"] = [];
          }
          process["bpmn:sequenceFlow"].push(sequenceFlow);

          const edge = {
            "@_id": `SequenceFlowEdge_${this._generateUniqueId()}`,
            "@_bpmnElement": sequenceFlow["@_id"],
            "di:waypoint": [
              { "@_x": Math.round((shape.BeginX || 0) * scaleFactor), "@_y": Math.round((shape.BeginY || 0) * scaleFactor) },
              { "@_x": Math.round((shape.EndX || 100) * scaleFactor), "@_y": Math.round((shape.EndY || 100) * scaleFactor) },
            ],
          };

          if (!bpmndiPlane["bpmndi:BPMNEdge"]) {
            bpmndiPlane["bpmndi:BPMNEdge"] = [];
          }
          bpmndiPlane["bpmndi:BPMNEdge"].push(edge);
        });
      }
    });
  }

  private _generateUniqueId(): string {
    return Math.random().toString(36).substr(2, 9);
  }

  private _createDefaultParticipant(process: any, bpmndiPlane: any, scaleFactor: number) {
    const participantElement = {
      "@_id": this.participantId,
      "@_name": "Default Pool",
      "@_processRef": this.processId,
    };

    // Update the collaboration to include the participant
    if (!this.bpmnJson["bpmn:definitions"]["bpmn:collaboration"]["bpmn:participant"]) {
      this.bpmnJson["bpmn:definitions"]["bpmn:collaboration"]["bpmn:participant"] = participantElement;
    } else if (Array.isArray(this.bpmnJson["bpmn:definitions"]["bpmn:collaboration"]["bpmn:participant"])) {
      this.bpmnJson["bpmn:definitions"]["bpmn:collaboration"]["bpmn:participant"].push(participantElement);
    } else {
      this.bpmnJson["bpmn:definitions"]["bpmn:collaboration"]["bpmn:participant"] = [this.bpmnJson["bpmn:definitions"]["bpmn:collaboration"]["bpmn:participant"], participantElement];
    }

    const participantShape = {
      "@_id": `BPMNShape_${this._generateUniqueId()}`,
      "@_bpmnElement": this.participantId,
      "dc:Bounds": {
        "@_x": "0",
        "@_y": "0",
        "@_width": "600",
        "@_height": "400",
      },
    };

    if (!bpmndiPlane["bpmndi:BPMNShape"]) {
      bpmndiPlane["bpmndi:BPMNShape"] = participantShape;
    } else if (Array.isArray(bpmndiPlane["bpmndi:BPMNShape"])) {
      bpmndiPlane["bpmndi:BPMNShape"].push(participantShape);
    } else {
      bpmndiPlane["bpmndi:BPMNShape"] = [bpmndiPlane["bpmndi:BPMNShape"], participantShape];
    }
  }

  private _createParticipant(process: any, bpmndiPlane: any, poolShape: any, scaleFactor: number) {
    const participantElement = {
      "@_id": this.participantId,
      "@_name": this._getShapeText(poolShape),
      "@_processRef": this.processId,
    };

    // Update the collaboration to include the participant
    if (!this.bpmnJson["bpmn:definitions"]["bpmn:collaboration"]["bpmn:participant"]) {
      this.bpmnJson["bpmn:definitions"]["bpmn:collaboration"]["bpmn:participant"] = participantElement;
    } else if (Array.isArray(this.bpmnJson["bpmn:definitions"]["bpmn:collaboration"]["bpmn:participant"])) {
      this.bpmnJson["bpmn:definitions"]["bpmn:collaboration"]["bpmn:participant"].push(participantElement);
    } else {
      this.bpmnJson["bpmn:definitions"]["bpmn:collaboration"]["bpmn:participant"] = [this.bpmnJson["bpmn:definitions"]["bpmn:collaboration"]["bpmn:participant"], participantElement];
    }

    const bounds = this._getShapeBounds(poolShape, scaleFactor);
    const participantShape = {
      "@_id": `BPMNShape_${this._generateUniqueId()}`,
      "@_bpmnElement": this.participantId,
      "dc:Bounds": bounds,
    };

    if (!bpmndiPlane["bpmndi:BPMNShape"]) {
      bpmndiPlane["bpmndi:BPMNShape"] = participantShape;
    } else if (Array.isArray(bpmndiPlane["bpmndi:BPMNShape"])) {
      bpmndiPlane["bpmndi:BPMNShape"].push(participantShape);
    } else {
      bpmndiPlane["bpmndi:BPMNShape"] = [bpmndiPlane["bpmndi:BPMNShape"], participantShape];
    }
  }
}
