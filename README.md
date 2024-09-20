<div class="custom-html md-stream-desktop"><h1>Excel to Docx Auto-Filler Tool</h1>
<h2>项目简介</h2>
<p id=""><strong>Excel to Docx Auto-Filler Tool</strong> 是一个基于Python的自动化工具，旨在帮助用户将Excel表格中的数据自动填充到Word文档的模板中，并生成多个格式化的Word文档。该工具特别适用于需要批量生成合同、报告等文档的场景，能够显著提高工作效率和减少人为错误。</p>
<h2>功能特点</h2>
<ul>
<li><strong>批量数据填充</strong>：支持从Excel文件中读取数据，并自动填充到Word文档的指定位置。</li>
<li><strong>自定义模板</strong>：用户可以根据需要自定义Word模板，包括文本格式、表格结构、图片位置等。</li>
<li><strong>图片处理</strong>：自动识别Excel中的图片链接，并下载图片插入到Word文档中。</li>
<li><strong>灵活的文件名模板</strong>：支持使用文件名模板，根据Excel中的数据动态生成文档名。</li>
<li><strong>图形用户界面</strong>：提供了基于Tkinter的图形用户界面，方便用户操作。</li>
</ul>
<h2>安装与依赖</h2>
<p id="">确保你的环境中已安装Python 3.x（推荐Python 3.11或更高版本），并安装以下依赖包：</p>
<pre><code class="language-bash"><div class="code-header"><span class="code-lang">bash</span><span class="code-copy" style=""><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 32 32"><path d="M28 1.333H9.333C8.597 1.333 8 1.93 8 2.667v4H4c-.736 0-1.333.597-1.333 1.333v14.667c0 .353.14.692.39.943l6.667 6.666c.25.25.589.39.943.39h12c.736 0 1.333-.596 1.333-1.333v-4h4c.736 0 1.333-.597 1.333-1.333V2.667c0-.737-.597-1.334-1.333-1.334zM9.333 26.115L7.22 24h2.114v2.115zm12 1.885H12v-5.333c0-.737-.597-1.334-1.333-1.334H5.333v-12h16V28zm5.334-5.333H24V8c0-.736-.597-1.333-1.333-1.333h-12V4h16v18.667z"></path></svg><span class="code-copy-text">复制代码</span></span></div><div class="code-wrapper"><table class="hljs hljs-ln"><tbody><tr><td class="hljs-ln-line hljs-ln-numbers" data-line-number="1"><div class="hljs-ln-n" data-line-number="1"></div></td><td class="hljs-ln-line hljs-ln-code" data-line-number="1">pip install pandas python-docx requests tkinter openpyxl</td></tr></tbody></table></div></code></pre>
<p id="">注意：<code class=" inline">tkinter</code>通常是Python标准库的一部分，无需单独安装。</p>
<h2>使用方法</h2>
<ol>
<li><strong>启动GUI</strong>：运行项目中的主Python脚本（通常是<code class=" inline">main.py</code>），将启动图形用户界面。</li>
<li><strong>配置参数</strong>：
<ul>
<li><strong>Excel文件路径</strong>：指定包含数据的Excel文件。</li>
<li><strong>DOCX模板路径</strong>：指定Word文档的模板文件。</li>
<li><strong>输出文件夹路径</strong>：设置生成的Word文档保存的文件夹。</li>
<li><strong>文件名模板</strong>：使用<code class=" inline">{{key}}</code>作为占位符，指定如何根据Excel数据生成文件名。</li>
<li><strong>图片尺寸</strong>：设置插入图片的宽度（英寸）。</li>
</ul>
</li>
<li><strong>选择文件与文件夹</strong>：使用提供的按钮选择文件或文件夹。</li>
<li><strong>生成文档</strong>：点击“生成docx文档”按钮，程序将自动处理并生成文档。</li>
</ol>
<h2>注意事项</h2>
<ul>
<li>确保Excel文件和Word模板中的占位符（如<code class=" inline">{{key}}</code>）与DataFrame中的列名完全匹配。</li>
<li>图片链接必须是可以直接访问的URL，并且指向的图片格式为jpg、jpeg、png或gif之一。</li>
<li>生成的图片文件会暂时保存在本地，并在文档生成后被删除。</li>
</ul>
<h2>贡献与反馈</h2>
<p id="">如果你在使用中发现任何问题或有改进建议，欢迎通过GitHub的Issue系统提交。我们也欢迎贡献代码或文档，以帮助项目不断完善。</p>
<h2>许可证</h2>
<p id="">本项目采用<a href="https://opensource.org/licenses/MIT" target="_blank" rel="noreferrer">MIT License</a>进行许可。</p>
<h2>开发者信息</h2>
<ul>
<li><strong>项目作者</strong>：钚尔什维克</li>
<li><strong>GitHub仓库</strong>：[[GitHub仓库链接]](https://github.com/Puershiweilke)</li>
<li><strong>个人网站</strong>：https://newhouseofme.top</li>
<li><strong>bilibili</strong>：https://space.bilibili.com/10720296</li>
<li><strong>抖音号</strong>：puershiweike</li>
</ul>
<h2>致谢</h2>
<p id="">将感谢所有为该项目提供代码、建议和反馈的贡献者。</p>

<div class="custom-html md-stream-desktop"><h1>Excel to Docx Auto-Filler Tool</h1>
<h2>project introduction</h2>
<p id=""><strong>Excel to Docx Auto-Filler Tool</strong> An automated tool based on Python, it aims to assist users in automatically filling Word document templates with data from Excel spreadsheets and generating multiple formatted Word documents. This tool is particularly suitable for scenarios requiring bulk generation of contracts, reports, and other documents, significantly improving work efficiency and reducing human errors.</p>
<h2>function features</h2>
<ul>
<li><strong>filling of batch data</strong>：支持从Excel文件中读取数据，并自动填充到Word文档的指定位置。</li>
<li><strong>custom editable template</strong>：用户可以根据需要自定义Word模板，包括文本格式、表格结构、图片位置等。</li>
<li><strong>图片处理</strong>：自动识别Excel中的图片链接，并下载图片插入到Word文档中。</li>
<li><strong>灵活的文件名模板</strong>：支持使用文件名模板，根据Excel中的数据动态生成文档名。</li>
<li><strong>图形用户界面</strong>：提供了基于Tkinter的图形用户界面，方便用户操作。</li>
</ul>
<h2>安装与依赖</h2>
<p id="">确保你的环境中已安装Python 3.x（推荐Python 3.11或更高版本），并安装以下依赖包：</p>
<pre><code class="language-bash"><div class="code-header"><span class="code-lang">bash</span><span class="code-copy" style=""><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 32 32"><path d="M28 1.333H9.333C8.597 1.333 8 1.93 8 2.667v4H4c-.736 0-1.333.597-1.333 1.333v14.667c0 .353.14.692.39.943l6.667 6.666c.25.25.589.39.943.39h12c.736 0 1.333-.596 1.333-1.333v-4h4c.736 0 1.333-.597 1.333-1.333V2.667c0-.737-.597-1.334-1.333-1.334zM9.333 26.115L7.22 24h2.114v2.115zm12 1.885H12v-5.333c0-.737-.597-1.334-1.333-1.334H5.333v-12h16V28zm5.334-5.333H24V8c0-.736-.597-1.333-1.333-1.333h-12V4h16v18.667z"></path></svg><span class="code-copy-text">复制代码</span></span></div><div class="code-wrapper"><table class="hljs hljs-ln"><tbody><tr><td class="hljs-ln-line hljs-ln-numbers" data-line-number="1"><div class="hljs-ln-n" data-line-number="1"></div></td><td class="hljs-ln-line hljs-ln-code" data-line-number="1">pip install pandas python-docx requests tkinter openpyxl</td></tr></tbody></table></div></code></pre>
<p id="">注意：<code class=" inline">tkinter</code>通常是Python标准库的一部分，无需单独安装。</p>
<h2>使用方法</h2>
<ol>
<li><strong>启动GUI</strong>：运行项目中的主Python脚本（通常是<code class=" inline">main.py</code>），将启动图形用户界面。</li>
<li><strong>配置参数</strong>：
<ul>
<li><strong>Excel文件路径</strong>：指定包含数据的Excel文件。</li>
<li><strong>DOCX模板路径</strong>：指定Word文档的模板文件。</li>
<li><strong>输出文件夹路径</strong>：设置生成的Word文档保存的文件夹。</li>
<li><strong>文件名模板</strong>：使用<code class=" inline">{{key}}</code>作为占位符，指定如何根据Excel数据生成文件名。</li>
<li><strong>图片尺寸</strong>：设置插入图片的宽度（英寸）。</li>
</ul>
</li>
<li><strong>选择文件与文件夹</strong>：使用提供的按钮选择文件或文件夹。</li>
<li><strong>生成文档</strong>：点击“生成docx文档”按钮，程序将自动处理并生成文档。</li>
</ol>
<h2>注意事项</h2>
<ul>
<li>确保Excel文件和Word模板中的占位符（如<code class=" inline">{{key}}</code>）与DataFrame中的列名完全匹配。</li>
<li>图片链接必须是可以直接访问的URL，并且指向的图片格式为jpg、jpeg、png或gif之一。</li>
<li>生成的图片文件会暂时保存在本地，并在文档生成后被删除。</li>
</ul>
<h2>贡献与反馈</h2>
<p id="">如果你在使用中发现任何问题或有改进建议，欢迎通过GitHub的Issue系统提交。我们也欢迎贡献代码或文档，以帮助项目不断完善。</p>
<h2>许可证</h2>
<p id="">本项目采用<a href="https://opensource.org/licenses/MIT" target="_blank" rel="noreferrer">MIT License</a>进行许可。</p>
<h2>开发者信息</h2>
<ul>
<li><strong>项目作者</strong>：钚尔什维克</li>
<li><strong>GitHub仓库</strong>：[[GitHub仓库链接]](https://github.com/Puershiweilke)</li>
<li><strong>个人网站</strong>：https://newhouseofme.top</li>
<li><strong>bilibili</strong>：https://space.bilibili.com/10720296</li>
<li><strong>抖音号</strong>：puershiweike</li>
</ul>
<h2>致谢</h2>
<p id="">将感谢所有为该项目提供代码、建议和反馈的贡献者。</p>
