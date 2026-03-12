import { useEffect, useRef, useState } from "react";

const AUTO_BLOCK_FORMULA_LENGTH = 24;
const LATEX_PATTERN =
  /(\$\$[\s\S]+?\$\$)|((^|[^\\])\$[^$\n]+?\$)|(\\\([\s\S]+?\\\))|(\\\[[\s\S]+?\\\])|(\\(?:frac|sqrt|sum|int|left|right|begin|end|alpha|beta|gamma|delta|theta|lambda|pi|times|cdot|pm|leq|geq|neq|ce|pu)\b)/;

function getPureInlineFormulaBody(value: string): string | null {
  const trimmed = value.trim();

  if (trimmed.startsWith("$$") && trimmed.endsWith("$$")) {
    return null;
  }
  if (trimmed.startsWith("\\[") && trimmed.endsWith("\\]")) {
    return null;
  }

  if (trimmed.startsWith("$") && trimmed.endsWith("$")) {
    const body = trimmed.slice(1, -1);
    if (!body.includes("$")) {
      return body.trim();
    }
  }

  if (trimmed.startsWith("\\(") && trimmed.endsWith("\\)")) {
    return trimmed.slice(2, -2).trim();
  }

  return null;
}

export function shouldAutoDisplayLatex(value: string): boolean {
  const formulaBody = getPureInlineFormulaBody(value);
  return (
    formulaBody !== null && formulaBody.length >= AUTO_BLOCK_FORMULA_LENGTH
  );
}

function toDisplayMathIfNeeded(value: string, forceDisplay: boolean): string {
  if (!forceDisplay) {
    return value;
  }
  const formulaBody = getPureInlineFormulaBody(value);
  if (formulaBody === null) {
    return value;
  }
  return `\\[${formulaBody}\\]`;
}

export function hasLatexSyntax(value: string): boolean {
  return LATEX_PATTERN.test(value.trim());
}

function hasMathDelimiter(value: string): boolean {
  const trimmed = value.trim();
  return (
    (trimmed.startsWith("$$") && trimmed.endsWith("$$")) ||
    (trimmed.startsWith("\\[") && trimmed.endsWith("\\]")) ||
    (trimmed.startsWith("$") && trimmed.endsWith("$")) ||
    (trimmed.startsWith("\\(") && trimmed.endsWith("\\)"))
  );
}

function toMathJaxSource(value: string, forceDisplay: boolean): string {
  const normalized = toDisplayMathIfNeeded(value, forceDisplay);
  if (hasMathDelimiter(normalized)) {
    return normalized;
  }

  const trimmed = normalized.trim();
  if (trimmed.length > 0 && hasLatexSyntax(trimmed)) {
    return `\\(${trimmed}\\)`;
  }

  return normalized;
}

type MathJaxConfig = {
  loader?: {
    load?: string[];
  };
  tex?: {
    inlineMath?: Array<[string, string]>;
    displayMath?: Array<[string, string]>;
    packages?: Record<string, string[]>;
  };
  options?: {
    skipHtmlTags?: string[];
  };
  svg?: {
    fontCache?: string;
  };
  startup?: {
    promise?: Promise<unknown>;
  };
  typesetPromise?: (elements?: Element[]) => Promise<unknown>;
};

declare global {
  interface Window {
    MathJax?: MathJaxConfig;
  }
}

let mathJaxLoadPromise: Promise<void> | null = null;

async function ensureMathJaxLoaded(): Promise<MathJaxConfig | null> {
  if (typeof window === "undefined") {
    return null;
  }

  if (window.MathJax?.typesetPromise) {
    return window.MathJax;
  }

  if (!mathJaxLoadPromise) {
    mathJaxLoadPromise = new Promise<void>((resolve, reject) => {
      const existingScript = document.querySelector<HTMLScriptElement>(
        "script[data-mathjax-loader='true']",
      );

      if (existingScript) {
        if (window.MathJax?.typesetPromise) {
          resolve();
          return;
        }
        existingScript.addEventListener("load", () => resolve(), {
          once: true,
        });
        existingScript.addEventListener(
          "error",
          () => reject(new Error("MathJax 脚本加载失败")),
          { once: true },
        );
        return;
      }

      window.MathJax = window.MathJax ?? {};
      const currentLoader = window.MathJax.loader ?? {};
      const currentLoaderLoads = currentLoader.load ?? [];
      window.MathJax.loader = {
        ...currentLoader,
        load: Array.from(new Set([...currentLoaderLoads, "[tex]/mhchem"])),
      };

      const currentTex = window.MathJax.tex ?? {};
      const currentPackages = currentTex.packages ?? {};
      const extraPackages = currentPackages["[+]"] ?? [];
      window.MathJax.tex = {
        ...currentTex,
        inlineMath: currentTex.inlineMath ?? [
          ["$", "$"],
          ["\\(", "\\)"],
        ],
        displayMath: currentTex.displayMath ?? [
          ["$$", "$$"],
          ["\\[", "\\]"],
        ],
        packages: {
          ...currentPackages,
          "[+]": Array.from(new Set([...extraPackages, "mhchem"])),
        },
      };
      window.MathJax.options = window.MathJax.options ?? {
        skipHtmlTags: [
          "script",
          "noscript",
          "style",
          "textarea",
          "pre",
          "code",
        ],
      };
      window.MathJax.svg = window.MathJax.svg ?? {
        fontCache: "global",
      };

      const script = document.createElement("script");
      script.src = "https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-svg.js";
      script.async = true;
      script.defer = true;
      script.setAttribute("data-mathjax-loader", "true");
      script.onload = () => resolve();
      script.onerror = () => reject(new Error("MathJax 脚本加载失败"));
      document.head.appendChild(script);
    }).catch((error) => {
      mathJaxLoadPromise = null;
      throw error;
    });
  }

  await mathJaxLoadPromise;
  return window.MathJax ?? null;
}

export function LatexRenderer({
  value,
  forceDisplay = false,
}: {
  value: string;
  forceDisplay?: boolean;
}) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [renderFailed, setRenderFailed] = useState(false);

  useEffect(() => {
    let disposed = false;

    const render = async () => {
      const container = containerRef.current;
      if (!container) {
        return;
      }

      container.textContent = toMathJaxSource(value, forceDisplay);
      setRenderFailed(false);

      try {
        const mathJax = await ensureMathJaxLoaded();
        if (disposed || !mathJax?.typesetPromise || !containerRef.current) {
          return;
        }

        if (mathJax.startup?.promise) {
          await mathJax.startup.promise;
        }
        await mathJax.typesetPromise([containerRef.current]);
      } catch {
        if (!disposed) {
          setRenderFailed(true);
        }
      }
    };

    void render();
    return () => {
      disposed = true;
    };
  }, [value, forceDisplay]);

  if (renderFailed) {
    return <div className="latex-plain">{value}</div>;
  }

  return (
    <div
      className={`latex-rendered ${forceDisplay ? "latex-rendered-display" : "latex-rendered-inline"}`}
      ref={containerRef}
    />
  );
}
