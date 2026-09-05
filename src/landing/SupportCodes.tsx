import { useEffect, useRef, useState } from "react";

type Code = { file: string; title: string; src: string };

export function SupportCodes() {
  const section = useRef<HTMLDivElement>(null);
  const [codes, setCodes] = useState<Code[]>([]);
  const [failed, setFailed] = useState(false);

  useEffect(() => {
    const controller = new AbortController();
    const load = () => {
      void fetch("/community/qr-codes.json", { signal: controller.signal })
        .then(response => {
          if (!response.ok) throw new Error("Image request failed");
          return response.json() as Promise<{ codes: Code[] }>;
        })
        .then(data => setCodes(data.codes))
        .catch(() => { if (!controller.signal.aborted) setFailed(true); });
    };
    const observer = new IntersectionObserver(entries => {
      if (entries.some(entry => entry.isIntersecting)) {
        observer.disconnect();
        load();
      }
    }, { rootMargin: "400px" });
    if (section.current) observer.observe(section.current);
    return () => { observer.disconnect(); controller.abort(); };
  }, []);

  return (
    <div ref={section}>
      <div className="landing-support-grid" aria-label="微信、支付宝与 QQ 群二维码">
        {codes.map(code => (
          <figure key={code.file}>
            <img src={code.src} alt={code.title + "二维码"} height={320} decoding="async" />
            <figcaption>{code.title}</figcaption>
          </figure>
        ))}
      </div>
      {failed && <p>二维码暂时未能加载。<a className="landing-text-link" href="https://github.com/paulhe666/visualtex#支持-visualtex" target="_blank" rel="noreferrer">在 GitHub 查看</a></p>}
    </div>
  );
}
