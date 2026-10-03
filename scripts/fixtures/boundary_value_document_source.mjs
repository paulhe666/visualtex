export const boundaryValueDocumentSource = String.raw`非齐次边界条件不容易给定本征基矢量，所以我们要先利用齐次边界条件来解本征值，然后再利用非齐次边界条件来进行系数叠加

如果我的边界条件没有齐次项，我们可以把这个方程拆分成两个方程问题：两个分别有一边是其次的。

非齐次边界条件破坏了边界的分离。

对于分离变量法的理解：对于齐次条件，我们的分离变量法是能够最容易找到一组本征基矢量，然后利用这个本征基矢量去做内积展开，因而得到比较好的解。对于非齐次条件，我们可以进行对非齐次项也用本征基矢量展开，得到每一个模式对应的方程。而且对于本征基矢量的选取本身是可以任意选的，我们可以根据方程的形式选取一个尽可能简单的形式来进行分解，减少耦合。

实际上，对于全是齐次边界条件的问题，如果是多维的，我们可以直接写出一组正交多维基矢量，然后用这个多维基矢量对方程的非齐次项进行展开。

\begin{equation*}
\frac{\partial^{2}u}{\partial x^{2}}+\frac{\partial^{2}u}{\partial y^{2}}=f(x,y),\qquad 0<x<a,\quad 0<y<b;
\end{equation*}

\begin{equation*}
\left.u\right|_{x=0}=0,\qquad \left.u\right|_{x=a}=0,\qquad 0\le y\le b,
\end{equation*}

\begin{equation*}
\left.u\right|_{y=0}=0,\qquad \left.u\right|_{y=b}=0,\qquad 0\le x\le a.
\end{equation*}可以展开成\begin{equation*}
u(x,y)=\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}c_{nm}\sin\frac{n\pi}{a}x\sin\frac{m\pi}{b}y,\qquad
f(x,y)=\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}d_{nm}\sin\frac{n\pi}{a}x\sin\frac{m\pi}{b}y.
\end{equation*}
对于边界条件全是非齐次的，我们可以这样子进行：把解函数分成两个函数的叠加，然后其中一个满足非齐次边界条件，一个满足齐次边界条件，然后非齐次边界条件的函数可以尽可能满足条件且简单
\begin{equation*}
\frac{\partial^{2}w}{\partial t^{2}}-a^{2}\frac{\partial^{2}w}{\partial x^{2}}
=-\left(\frac{\partial^{2}v}{\partial t^{2}}-a^{2}\frac{\partial^{2}v}{\partial x^{2}}\right),
\qquad 0<x<l,\ t>0;
\end{equation*}

\begin{equation*}
\left.w\right|_{t=0}=-\left.v\right|_{t=0},\qquad
\left.\frac{\partial u}{\partial t}\right|_{t=0}
=-\left.\frac{\partial v}{\partial t}\right|_{t=0},
\qquad 0\le x\le l.
\end{equation*}
对于一个微分算子，如果我们解出来他的本征函数要么是奇函数，要么是偶函数，则我们知道它具有某种反射对称性。

利用对易子可以对正交曲面方程问题进行很好的分离变量。例如，在球坐标、柱坐标等进行分离变量用的都是对易子，各个方向的算子都是具有相同本征基矢的。

力学量完全集，集合使得指定确定的一个本征值，有唯一一个本征函数与之对应

\chapter{正交曲面坐标系}
\section{度规矩阵}
如果度规矩阵是对角矩阵，那么这个坐标系是正交曲面坐标系，因为这样的话新坐标基矢之间就是相互正交的，即对角元为零

laplace算子的表达式：\[
g_{ii}=h_i^2,\qquad \sqrt{g}=h_1h_2h_3
\]

\[
\nabla^2=\frac{1}{h_1h_2h_3}\sum_{i=1}^3\frac{\partial}{\partial x^i}\left(\frac{h_1h_2h_3}{h_i^2}\frac{\partial}{\partial x^i}\right)
\]

\[
\nabla^2=\frac{1}{h_1h_2h_3}\left[
\frac{\partial}{\partial x^1}\left(\frac{h_2h_3}{h_1}\frac{\partial}{\partial x^1}\right)
+\frac{\partial}{\partial x^2}\left(\frac{h_1h_3}{h_2}\frac{\partial}{\partial x^2}\right)
+\frac{\partial}{\partial x^3}\left(\frac{h_1h_2}{h_3}\frac{\partial}{\partial x^3}\right)
\right]
\]
在圆形边界中，本征值只可取离散值，原因可以理解为是因为在周期边界中出现自干涉，导致的量子化条件，和玻尔-索末菲的量子化条件有点像

图伦不变量？是什么

圆形边界内laplace方程的第一类边值问题的通解可以表示为：平面上的电势多极展开
规定原点处有界可以避免多解的情况

利用圆形边界的解的系数形式，可以写成：
\[
u(r,\varphi)=\frac{a^{2}-r^{2}}{2\pi}\int_{0}^{2\pi}\frac{g(\varphi')}{r^{2}+a^{2}-2ar\cos(\varphi-\varphi')}\,d\varphi'
\]

柱贝塞尔方程：\[
\frac{1}{r}\frac{d}{dr}\left(r\frac{dR}{dr}\right)+\left[k^{2}-\lambda-\frac{\mu}{r^{2}}\right]R=0.
\]
球坐标分离变量

legendre方程也是一个SL方程，而且如果$sin\theta$在边界为零，则这个微分算子是自伴的。有奇异性的SL方程？

角向的两个方程一个就是$L^2$和$L_z$，而且它们是对易的。所以他们构成共同本征基。

然后$L_z=-i\hbar\frac{\partial}{\partial \varphi}$,理解方法是我们类比动量的表达式，可以凑出来这个形式，我们也可以直接算出来：由
\begin{equation*}
L=-i\hbar \vec{r} \times \nabla
\end{equation*}
所以
\[
\hat{L}_z = x \hat{p}_y - y \hat{p}_x = - i \hbar \left( x \frac{\partial}{\partial y} - y \frac{\partial}{\partial x} \right)
\]
然后把直角坐标系变成球坐标系即可

\begin{equation*}
L=-\ i\hbar \left(\overrightarrow{e_{\varphi }}\frac{\partial }{\partial \theta } -\overrightarrow{e_{\theta }}\frac{1}{sin\theta }\frac{\partial }{\partial \varphi }\right)
\end{equation*}
坐标变换的时候要注意奇异项的问题，会多一个delta函数的解，所以要引入自然边界条件来解决这个问题
\begin{equation*}
L_{z} =-i\hbar \frac{\partial }{\partial \varphi }
\end{equation*}
角动量算子和旋转对称变换
对于一个x平移变换，我们可以写成
\begin{equation*}
\mathcal{T}( a) =e^{-\frac{i}{\hbar } aP_{x}}
\end{equation*}
一个算符造成物理事实不变的表述是：这个算符和哈密顿量对易，然后根据平移算符导致哈密顿量不变，所以推出动量算符和哈密顿量对易。

角动量算符：
\[
R_{n}(\theta)=\exp\left[-\frac{i}{\hbar}\theta L_{n}\right]
\]
一种更深层次的理解：角动量算子是旋转变换的生成元，然后由角动量算子的不对易性和乘法逆的封闭性，可以自然得到\[
[L_x,L_y]=iL_z
\]，所以实际上只要满足这些东西的算子都可以被叫做角动量算子`;
