import canvas from "canvas";
import { Window } from "happy-dom";

let shimInstalled = false;
let windowInstance: InstanceType<typeof Window> | null = null;

type EventHandler = ((event: { target: NodeFileReader }) => void) | null;

class NodeFileReader {
  result: string | ArrayBuffer | null = null;
  error: unknown = null;
  onload: EventHandler = null;
  onerror: EventHandler = null;
  onloadend: EventHandler = null;
  readyState = 0;

  private done(result: string | ArrayBuffer): void {
    this.readyState = 2;
    this.result = result;
    const event = { target: this };
    this.onload?.(event);
    this.onloadend?.(event);
  }

  private fail(error: unknown): void {
    this.readyState = 2;
    this.error = error;
    const event = { target: this };
    this.onerror?.(event);
    this.onloadend?.(event);
  }

  readAsArrayBuffer(input: unknown): void {
    this.readyState = 1;

    if (input instanceof ArrayBuffer) {
      process.nextTick(() => this.done(input));
      return;
    }

    if (isArrayBufferSource(input)) {
      input
        .arrayBuffer()
        .then((bytes) => this.done(bytes))
        .catch((error) => this.fail(error));
    }
  }

  readAsBinaryString(input: unknown): void {
    this.readyState = 1;

    if (!isArrayBufferSource(input)) {
      return;
    }

    input
      .arrayBuffer()
      .then((bytes) => {
        const data = new Uint8Array(bytes);
        let text = "";
        for (const value of data) {
          text += String.fromCharCode(value);
        }
        this.done(text);
      })
      .catch((error) => this.fail(error));
  }

  readAsDataURL(input: unknown): void {
    this.readyState = 1;

    if (!isArrayBufferSource(input)) {
      return;
    }

    input
      .arrayBuffer()
      .then((bytes) => {
        const mimeType = input.type ?? "application/octet-stream";
        const base64 = Buffer.from(bytes).toString("base64");
        this.done(`data:${mimeType};base64,${base64}`);
      })
      .catch((error) => this.fail(error));
  }

  readAsText(input: unknown, encoding?: string): void {
    this.readyState = 1;

    if (!isArrayBufferSource(input)) {
      return;
    }

    input
      .arrayBuffer()
      .then((bytes) => {
        const decoder = new TextDecoder(encoding ?? "utf-8");
        this.done(decoder.decode(bytes));
      })
      .catch((error) => this.fail(error));
  }

  abort(): void {}

  addEventListener(
    type: "load" | "error" | "loadend",
    listener: EventHandler,
  ): void {
    if (!listener) {
      return;
    }
    if (type === "load") {
      this.onload = listener;
    } else if (type === "error") {
      this.onerror = listener;
    } else {
      this.onloadend = listener;
    }
  }

  removeEventListener(
    type: "load" | "error" | "loadend",
    listener: EventHandler,
  ): void {
    if (type === "load" && this.onload === listener) {
      this.onload = null;
    }
    if (type === "error" && this.onerror === listener) {
      this.onerror = null;
    }
    if (type === "loadend" && this.onloadend === listener) {
      this.onloadend = null;
    }
  }
}

interface ArrayBufferSourceLike {
  arrayBuffer: () => Promise<ArrayBuffer>;
  type?: string;
}

function isArrayBufferSource(value: unknown): value is ArrayBufferSourceLike {
  return (
    typeof value === "object" &&
    value !== null &&
    "arrayBuffer" in value &&
    typeof (value as { arrayBuffer?: unknown }).arrayBuffer === "function"
  );
}

function setGlobal(name: string, value: unknown): void {
  try {
    Object.assign(globalThis, { [name]: value });
  } catch {
    Object.defineProperty(globalThis, name, {
      value,
      writable: true,
      configurable: true,
    });
  }
}

export function installShims(): void {
  if (shimInstalled) {
    return;
  }

  const win = new Window({ url: "http://localhost" });
  windowInstance = win;

  setGlobal("self", globalThis);
  setGlobal("window", win);
  setGlobal("document", win.document);
  setGlobal("navigator", win.navigator);
  setGlobal("HTMLCollection", win.HTMLCollection);
  setGlobal("getComputedStyle", win.getComputedStyle.bind(win));
  setGlobal("customElements", win.customElements);
  setGlobal("HTMLElement", win.HTMLElement);
  setGlobal("HTMLDivElement", win.HTMLDivElement);
  setGlobal("HTMLCanvasElement", win.HTMLCanvasElement);
  setGlobal("HTMLImageElement", win.HTMLImageElement);
  setGlobal("Image", win.Image);
  setGlobal("Event", win.Event);
  setGlobal("MouseEvent", win.MouseEvent);
  setGlobal("KeyboardEvent", win.KeyboardEvent);
  setGlobal("PointerEvent", win.PointerEvent || win.MouseEvent);
  setGlobal("TouchEvent", win.TouchEvent || class TouchEvent {});
  setGlobal("WheelEvent", win.WheelEvent || win.Event);
  setGlobal("MutationObserver", win.MutationObserver);
  setGlobal(
    "ResizeObserver",
    class ResizeObserver {
      observe(): void {}
      disconnect(): void {}
      unobserve(): void {}
    },
  );
  setGlobal("requestAnimationFrame", (callback: (timestamp: number) => void) =>
    setTimeout(() => callback(Date.now()), 0),
  );
  setGlobal("cancelAnimationFrame", (id: ReturnType<typeof setTimeout>) =>
    clearTimeout(id),
  );
  setGlobal("FileReader", NodeFileReader);
  setGlobal("DOMParser", win.DOMParser);
  setGlobal("XMLSerializer", win.XMLSerializer);
  setGlobal("canvas", canvas);
  setGlobal("devicePixelRatio", 1);
  setGlobal("location", win.location);
  setGlobal("innerWidth", 800);
  setGlobal("innerHeight", 600);
  setGlobal("addEventListener", () => {});
  setGlobal("removeEventListener", () => {});
  setGlobal("getSelection", () => ({
    removeAllRanges: () => {},
    addRange: () => {},
  }));

  shimInstalled = true;
}

export function disposeShims(): void {
  if (windowInstance) {
    windowInstance.close();
    windowInstance = null;
  }
  shimInstalled = false;
}

export function isShimInstalled(): boolean {
  return shimInstalled;
}
