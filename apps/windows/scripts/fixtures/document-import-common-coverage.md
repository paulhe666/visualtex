---
title: VisualTeX Markdown Coverage
author: VisualTeX
---

# VisualTeX Markdown Import

Setext level two
----------------

This paragraph contains **bold**, __strong__, *italic*, _emphasis_,
***bold italic***, ~~strike~~, `inline_code()`, and ``code with ` tick``.
It also contains the inline formulas $E=mc^2$ and \(a+b=c\), plus a hard break.<br>
This sentence starts after the hard break.

An inline link is [Example](https://example.com), a shortcut reference is
[OpenAI], an automatic link is <https://example.org>, and an image is
![Diagram](diagram.png).

[OpenAI]: https://openai.com

HTML entities remain readable: A &amp; B, &#x03B1; &lt; x &gt;.
<!-- This comment must not appear in Word. -->
A raw HTML wrapper keeps <strong>visible content</strong> and a break<br>after it.

> A block quotation with **formatting** and formula $q=1$.
> A second quoted line.

- First bullet with $x_1$
- [x] Completed task
- [ ] Pending task
  - Nested bullet

1. First numbered item
2) Second numbered item
    1. Nested numbered item

| Name | Value | Formula |
|:-----|------:|:--------|
| Alpha | 1 | $\alpha$ |
| Beta | 2 | $\beta$ |

---

```typescript
const formula = "$not_math$";
```

~~~python
print("tilde fence")
~~~

    indented_code = True

Display formulas:

$$
\int_0^1 x^2\,\mathrm{d}x=\frac13
$$

\[
\begin{aligned}
a+b&=c\\
d&=e+f
\end{aligned}
\]

Text with a footnote.[^note]

[^note]: This is the footnote body with $n=1$.
