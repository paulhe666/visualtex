import { useEffect, useRef, useState } from "react";
import { ART_TEXT, type InkKey } from "./art.generated";
import type { LandingLang } from "./i18n";

type Code = { file: string; title: string; src: string };

// Tile order and captions follow public/community/qr-codes.json.
const CAPTIONS: Record<string, InkKey> = {
  "wechat-pay.jpg": "qrWechat",
  "alipay.jpg": "qrAlipay",
  "qq-group.png": "qrGroup",
};
const ALT: Record<string, { zh: string; en: string }> = {
  "wechat-pay.jpg": { zh: "微信打赏二维码", en: "WeChat Pay QR code" },
  "alipay.jpg": { zh: "支付宝打赏二维码", en: "Alipay QR code" },
  "qq-group.png": { zh: "QQ 交流群二维码", en: "QQ group QR code" },
};
const FILES = Object.keys(CAPTIONS);

/** Three fixed-size tiles; the images are fetched once the section is near the viewport. */
export function SupportCodes({ lang }: { lang: LandingLang }) {
  const section = useRef<HTMLDivElement>(null);
  const [codes, setCodes] = useState<Record<string, string>>({});
  const [failed, setFailed] = useState(false);

  useEffect(() => {
    const controller = new AbortController();
    const load = () => {
      void fetch("/community/qr-codes.json", { signal: controller.signal })
        .then((response) => {
          if (!response.ok) throw new Error("Image request failed");
          return response.json() as Promise<{ codes: Code[] }>;
        })
        .then((data) => setCodes(Object.fromEntries(data.codes.map((code) => [code.file, code.src]))))
        .catch(() => { if (!controller.signal.aborted) setFailed(true); });
    };
    const observer = new IntersectionObserver((entries) => {
      if (entries.some((entry) => entry.isIntersecting)) {
        observer.disconnect();
        load();
      }
    }, { rootMargin: "400px" });
    if (section.current) observer.observe(section.current);
    return () => { observer.disconnect(); controller.abort(); };
  }, []);

  return (
    <div ref={section}>
      <div className="landing-qr">
        {FILES.map((file) => (
          <figure key={file}>
            <div className="landing-qr-img">
              {codes[file] && <img src={codes[file]} alt={ALT[file][lang]} decoding="async" />}
            </div>
            <figcaption className="landing-ink" dangerouslySetInnerHTML={{ __html: ART_TEXT[CAPTIONS[file]][lang] }} />
          </figure>
        ))}
      </div>
      {failed && (
        <p className="landing-qr-failed">
          <span className="landing-ink" dangerouslySetInnerHTML={{ __html: ART_TEXT.qrFailed[lang] }} />
          <a href="https://github.com/paulhe666/visualtex#支持-visualtex" target="_blank" rel="noreferrer">GitHub</a>
        </p>
      )}
    </div>
  );
}
