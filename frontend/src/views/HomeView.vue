<template>
  <main class="handwriting-workspace">
    <section class="workspace-shell">
      <aside class="control-column">
        <header class="workspace-header">
          <p>手写生成工作台</p>
          <h1>上传或编辑正文，一次完成预览与导出</h1>
        </header>

        <div v-if="message" class="notice notice-success">{{ message }}</div>
        <div v-if="uploadMessage" class="notice notice-info">{{ uploadMessage }}</div>
        <div v-if="errorMessage" class="notice notice-error">{{ errorMessage }}</div>

        <section class="panel source-panel">
          <div class="panel-heading">
            <span>正文来源</span>
            <small>PDF / Word / Markdown / TXT</small>
          </div>

          <div
            class="source-upload-target"
            :class="{ disabled: isExtractingSource }"
            @dragover.prevent
            @drop.prevent="onSourceFileDrop"
          >
            <input
              id="source-file-input"
              type="file"
              ref="sourceFileInput"
              accept=".pdf,.doc,.docx,.md,.markdown,.txt,.rtf"
              @change="onSourceFileChange"
              class="source-file-input"
              :disabled="isExtractingSource"
              aria-label="选择正文文件"
            />
            <div class="drop-zone" aria-hidden="true">
              <strong>{{ selectedSourceFileName || '选择正文文件' }}</strong>
              <span>{{ isExtractingSource ? '正在识别文档内容...' : '点击选择，或把文件拖到这里' }}</span>
            </div>
          </div>

          <div v-if="selectedSourceFileName" class="source-meta">
            <span>{{ selectedSourceFileName }}</span>
            <button type="button" @click="clearSourceFile">清空文件</button>
          </div>

          <textarea
            v-model="text"
            class="text-editor"
            placeholder="在这里输入或校对识别后的正文。支持 Markdown 和 $...$ 公式。"
          ></textarea>

          <div class="source-actions">
            <button
              type="button"
              @click="downloadStandardDocx"
              :disabled="!text.trim() || isPreparingStandardDocx"
            >
              {{ isPreparingStandardDocx ? '生成校对稿...' : '标准Word校对稿' }}
            </button>
          </div>
        </section>

        <section class="panel settings-panel">
          <div class="panel-heading">
            <span>样式设置</span>
            <small>保持原手写渲染效果</small>
          </div>

          <div class="settings-grid">
            <label>
              导出格式
              <select v-model="outputFormat">
                <option value="pdf">PDF</option>
                <option value="docx">Word</option>
              </select>
            </label>

            <label>
              字体
              <select v-model="selectedOption">
                <option v-for="option in options" :value="option.value" :key="option.value">
                  {{ option.text }}
                </option>
              </select>
            </label>

            <label>
              字号
              <input type="number" v-model.number="fontSize" />
            </label>

            <label>
              行距
              <input type="number" v-model.number="lineSpacing" />
            </label>

            <label>
              宽度
              <input type="number" v-model.number="width" :disabled="isBackgroundImageSpecified" />
            </label>

            <label>
              高度
              <input type="number" v-model.number="height" :disabled="isBackgroundImageSpecified" />
            </label>
          </div>

          <div class="asset-row">
            <input type="file" ref="fontFileInput" @change="onFontChange" accept=".ttf" hidden />
            <button type="button" @click="triggerFontFileInput">上传字体</button>
            <span>{{ selectedFontFileName || '使用字体列表中的字体' }}</span>
          </div>

          <div class="asset-row">
            <input type="file" ref="imageFileInput" @change="onBackgroundImageChange" accept=".png,.jpg,.jpeg" hidden />
            <button type="button" @click="triggerImageFileInput" :disabled="isDimensionSpecified">上传背景</button>
            <span>{{ selectedImageFileName || '未上传时自动生成横线背景' }}</span>
            <button v-if="selectedImageFileName" type="button" class="link-button" @click="clearImage">移除</button>
          </div>

          <div class="toggle-row">
            <label><input type="checkbox" v-model="isUnderlined" /> 增加下划线</label>
            <label><input type="checkbox" v-model="enableEnglishSpacing" /> 增大英文单词间距</label>
          </div>

          <details class="advanced-settings">
            <summary>高级扰动与边距</summary>
            <div class="settings-grid compact">
              <label>上边距<input type="number" v-model.number="marginTop" /></label>
              <label>下边距<input type="number" v-model.number="marginBottom" /></label>
              <label>左边距<input type="number" v-model.number="marginLeft" /></label>
              <label>右边距<input type="number" v-model.number="marginRight" /></label>
              <label>字间距<input type="number" v-model.number="wordSpacing" /></label>
              <label>行距扰动<input type="number" v-model.number="lineSpacingSigma" /></label>
              <label>字号扰动<input type="number" v-model.number="fontSizeSigma" /></label>
              <label>字距扰动<input type="number" v-model.number="wordSpacingSigma" /></label>
              <label>横向偏移<input type="number" v-model.number="perturbXSigma" /></label>
              <label>纵向偏移<input type="number" v-model.number="perturbYSigma" /></label>
              <label>旋转偏移<input type="number" v-model.number="perturbThetaSigma" step="0.01" /></label>
              <label>墨色扰动<input type="number" v-model.number="ink_depth_sigma" /></label>
              <label>涂改概率<input type="number" v-model.number="strikethrough_probability" step="0.001" /></label>
              <label>涂改长度<input type="number" v-model.number="strikethrough_length_sigma" /></label>
              <label>涂改宽度<input type="number" v-model.number="strikethrough_width" /></label>
              <label>涂改宽扰动<input type="number" v-model.number="strikethrough_width_sigma" /></label>
              <label>涂改角度<input type="number" v-model.number="strikethrough_angle_sigma" /></label>
            </div>
          </details>
        </section>

        <section class="actions-panel">
          <button type="button" class="secondary-action" @click="loadPreset">载入设置</button>
          <button type="button" class="secondary-action" @click="savePreset">保存设置</button>
          <button type="button" class="secondary-action" @click="resetSettings">重置</button>
          <button type="button" class="primary-action" @click="generateHandwriting(true)" :disabled="shouldDisableButtons">
            {{ isGenerating ? '生成中...' : '预览' }}
          </button>
          <button
            v-if="isDevEnv"
            type="button"
            class="secondary-action"
            @click="toggleFullPreview"
            :disabled="shouldDisableButtons"
          >
            全量预览 {{ enableFullPreview ? '开' : '关' }}
          </button>
          <button type="button" class="primary-action export" @click="generateHandwriting(false)" :disabled="shouldDisableButtons">
            {{ isGenerating ? '生成中...' : `导出 ${outputFormatLabel}` }}
          </button>
        </section>
      </aside>

      <section class="preview-column">
        <div class="preview-header">
          <div>
            <p>预览</p>
            <h2>{{ previewTitle }}</h2>
          </div>
          <div v-if="previewImages.length > 1" class="page-nav">
            <button type="button" @click="prevPage" :disabled="currentPreviewIndex === 0">上一页</button>
            <span>{{ currentPreviewIndex + 1 }} / {{ previewImages.length }}</span>
            <button type="button" @click="nextPage" :disabled="currentPreviewIndex === previewImages.length - 1">下一页</button>
          </div>
        </div>

        <div class="paper-preview">
          <img
            v-if="previewImages.length > 0"
            :src="previewImages[currentPreviewIndex]"
            :alt="`手写预览第 ${currentPreviewIndex + 1} 页`"
          />
          <img v-else :src="previewImage" alt="手写预览" />
        </div>
      </section>
    </section>

    <footer class="workspace-footer">
      <span>个人私有工具</span>
      <span>本地优先，按需部署</span>
    </footer>
  </main>
</template>

<script>
import Swal from 'sweetalert2';

const SETTINGS_KEYS = [
  'text',
  'fontSize',
  'lineSpacing',
  'fill',
  'width',
  'height',
  'marginTop',
  'marginBottom',
  'marginLeft',
  'marginRight',
  'selectedFontFileName',
  'selectedOption',
  'lineSpacingSigma',
  'fontSizeSigma',
  'wordSpacingSigma',
  'perturbXSigma',
  'perturbYSigma',
  'perturbThetaSigma',
  'wordSpacing',
  'strikethrough_length_sigma',
  'strikethrough_angle_sigma',
  'strikethrough_width_sigma',
  'strikethrough_probability',
  'strikethrough_width',
  'ink_depth_sigma',
  'isUnderlined',
  'enableEnglishSpacing',
  'outputFormat',
];

export default {
  name: 'HomeView',
  data() {
    return {
      text: '',
      sourceFile: null,
      selectedSourceFileName: '',
      sourceContentFormat: 'plain',
      isExtractingSource: false,
      fontFile: null,
      selectedFontFileName: '',
      backgroundImage: null,
      selectedImageFileName: '',
      options: [],
      selectedOption: '1',
      fontSize: 124,
      lineSpacing: 200,
      fill: '(0, 0, 0, 255)',
      width: 2481,
      height: 3507,
      marginTop: 50,
      marginBottom: 50,
      marginLeft: 50,
      marginRight: 50,
      lineSpacingSigma: 0,
      fontSizeSigma: 2,
      wordSpacingSigma: 2,
      perturbXSigma: 3,
      perturbYSigma: 3,
      perturbThetaSigma: 0.05,
      wordSpacing: 1,
      strikethrough_length_sigma: 2,
      strikethrough_angle_sigma: 2,
      strikethrough_width_sigma: 2,
      strikethrough_probability: 0.005,
      strikethrough_width: 8,
      ink_depth_sigma: 30,
      isUnderlined: true,
      enableEnglishSpacing: false,
      outputFormat: 'pdf',
      previewImage: '/default1.webp',
      previewImages: [],
      currentPreviewIndex: 0,
      message: '',
      uploadMessage: '',
      errorMessage: '',
      isLoading: false,
      isGenerating: false,
      isPreparingStandardDocx: false,
      lastGenerateTime: 0,
      generateCooldown: 3000,
      cooldownTimer: null,
      remainingCooldown: 0,
      isInCooldownPeriod: false,
      queueFullCountdown: 0,
      queueFullTotal: 0,
      queueFullTimer: null,
      enableFullPreview: false,
    };
  },
  computed: {
    isDimensionSpecified() {
      return !!(this.width || this.height);
    },
    isBackgroundImageSpecified() {
      return !!this.backgroundImage;
    },
    shouldDisableButtons() {
      return this.isGenerating || this.isInCooldownPeriod || this.queueFullCountdown > 0 || this.isExtractingSource;
    },
    isDevEnv() {
      return process.env.NODE_ENV === 'development';
    },
    outputFormatLabel() {
      return this.outputFormat === 'docx' ? 'Word' : 'PDF';
    },
    previewTitle() {
      if (this.previewImages.length > 0) {
        return `第 ${this.currentPreviewIndex + 1} 页`;
      }
      return '等待生成手写预览';
    },
  },
  watch: {
    text(value) {
      localStorage.setItem('text', JSON.stringify(value));
    },
    outputFormat(value) {
      localStorage.setItem('outputFormat', JSON.stringify(value));
    },
    errorMessage(value) {
      if (value) {
        this.$swal.fire({
          toast: true,
          position: 'top-end',
          icon: 'error',
          title: value,
          showConfirmButton: false,
          timer: 5000,
          timerProgressBar: true,
        });
      }
    },
    message(value) {
      if (value) {
        this.$swal.fire({
          toast: true,
          position: 'top-end',
          icon: 'success',
          title: value,
          showConfirmButton: false,
          timer: 3000,
          timerProgressBar: true,
        });
      }
    },
  },
  created() {
    SETTINGS_KEYS.forEach((item) => {
      const value = localStorage.getItem(item);
      if (value !== null && value !== 'undefined') {
        try {
          this[item] = JSON.parse(value);
        } catch (error) {
          localStorage.removeItem(item);
        }
      }
    });
    this.loadFonts();
  },
  beforeUnmount() {
    if (this.cooldownTimer) {
      clearInterval(this.cooldownTimer);
    }
    if (this.queueFullTimer) {
      clearInterval(this.queueFullTimer);
    }
  },
  methods: {
    async loadFonts() {
      try {
        const response = await this.$http.get('/api/fonts_info');
        this.options = response.data.map((font, index) => ({ value: String(index + 1), text: font }));
        if (!this.options.find((option) => option.value === this.selectedOption)) {
          this.selectedOption = this.options[0]?.value || '1';
        }
      } catch (error) {
        this.errorMessage = error.response?.data?.error || error.message || '字体列表加载失败';
      }
    },
    async onSourceFileChange(event) {
      const file = event.target.files[0];
      if (!file) return;
      await this.handleSourceFile(file, event);
    },
    async onSourceFileDrop(event) {
      if (this.isExtractingSource) return;
      const file = event.dataTransfer?.files?.[0];
      if (!file) return;
      await this.handleSourceFile(file);
    },
    async handleSourceFile(file, event = null) {
      const allowedSuffixes = ['.pdf', '.doc', '.docx', '.md', '.markdown', '.txt', '.rtf'];
      const lowerName = file.name.toLowerCase();
      if (!allowedSuffixes.some((suffix) => lowerName.endsWith(suffix))) {
        this.errorMessage = '只支持 PDF、Word、Markdown、TXT、RTF 文件';
        if (event?.target) {
          event.target.value = null;
        }
        return;
      }
      this.sourceFile = file;
      this.selectedSourceFileName = file.name;
      await this.extractSourceFile();
    },
    async extractSourceFile() {
      if (!this.sourceFile) return;
      this.isExtractingSource = true;
      this.message = '';
      this.errorMessage = '';
      this.uploadMessage = '正在识别文档内容...';
      const formData = new FormData();
      formData.append('file', this.sourceFile);
      try {
        const response = await this.$http.post('/api/handwriting/extract_source', formData, {
          headers: { 'Content-Type': 'multipart/form-data' },
          timeout: 10 * 60 * 1000,
        });
        this.text = response.data.markdown || '';
        this.sourceContentFormat = 'markdown';
        const warnings = response.data.warnings || [];
        this.message = warnings.length ? `识别完成，但有提示：${warnings.join('；')}` : '文档内容已识别并规范化，请校对后预览或导出';
        this.uploadMessage = '';
      } catch (error) {
        this.errorMessage = error.response?.data?.message || error.message || '文档识别失败';
        this.uploadMessage = '';
      } finally {
        this.isExtractingSource = false;
      }
    },
    clearSourceFile() {
      this.sourceFile = null;
      this.selectedSourceFileName = '';
      this.sourceContentFormat = this.detectMarkdownContent() ? 'markdown' : 'plain';
      if (this.$refs.sourceFileInput) {
        this.$refs.sourceFileInput.value = null;
      }
    },
    triggerFontFileInput() {
      this.$refs.fontFileInput.click();
    },
    onFontChange(event) {
      const file = event.target.files[0];
      if (!file) return;
      this.fontFile = file;
      this.selectedFontFileName = file.name;
      const newOption = { value: String(this.options.length + 1), text: file.name };
      this.options = [...this.options.filter((option) => option.text !== file.name), newOption];
      this.selectedOption = newOption.value;
      this.message = '字体已载入';
    },
    triggerImageFileInput() {
      if (!this.isDimensionSpecified) {
        this.$refs.imageFileInput.click();
      }
    },
    onBackgroundImageChange(event) {
      const file = event.target.files[0];
      if (!file) return;
      this.backgroundImage = file;
      this.selectedImageFileName = file.name;
      this.previewImage = URL.createObjectURL(file);
      Swal.fire({
        title: '是否自动识别页面边距？',
        icon: 'question',
        showCancelButton: true,
        confirmButtonText: '识别',
        cancelButtonText: '跳过',
      }).then((result) => {
        if (result.isConfirmed) {
          this.identifyBackgroundMargins();
        }
      });
    },
    async identifyBackgroundMargins() {
      const formData = new FormData();
      formData.append('file', this.backgroundImage);
      this.isLoading = true;
      try {
        const response = await this.$http.post('/api/imagefileprocess', formData, {
          headers: { 'Content-Type': 'multipart/form-data' },
        });
        this.marginLeft = response.data.marginLeft;
        this.marginRight = response.data.marginRight;
        this.marginTop = response.data.marginTop - this.lineSpacing;
        this.marginBottom = response.data.marginBottom;
        this.lineSpacing = response.data.lineSpacing;
        this.message = '背景图片已加载并识别边距';
      } catch (error) {
        this.errorMessage = error.response?.data?.error || error.message || '背景识别失败';
      } finally {
        this.isLoading = false;
      }
    },
    clearImage() {
      this.backgroundImage = null;
      this.selectedImageFileName = '';
      this.previewImage = '/default1.webp';
      if (this.$refs.imageFileInput) {
        this.$refs.imageFileInput.value = null;
      }
    },
    detectMarkdownContent() {
      return /(^|\n)\s{0,3}#{1,6}\s|\$[^$]+\$|\\(?:frac|sqrt|sum|int|begin)/.test(this.text || '');
    },
    activeContentFormat() {
      if (this.sourceContentFormat === 'markdown') return 'markdown';
      return this.detectMarkdownContent() ? 'markdown' : 'plain';
    },
    async downloadStandardDocx() {
      if (!this.text || !this.text.trim()) {
        this.errorMessage = '请先输入或上传正文';
        return;
      }
      this.isPreparingStandardDocx = true;
      this.message = '';
      this.errorMessage = '';
      this.uploadMessage = '正在生成标准Word校对稿...';
      const formData = new FormData();
      formData.append('markdown', this.text);
      formData.append('filename', `${this.selectedSourceFileName || 'standard_formula_draft'}.docx`);
      try {
        const response = await this.$http.post('/api/handwriting/markdown_docx', formData, {
          headers: { 'Content-Type': 'multipart/form-data' },
          responseType: 'blob',
          timeout: 120000,
        });
        const filename = this.filenameFromDisposition(response.headers['content-disposition'], 'standard_formula_draft.docx');
        this.downloadBlob(response.data, filename);
        this.message = `标准Word校对稿已导出：${filename}`;
        this.uploadMessage = '';
      } catch (error) {
        if (error.response?.data instanceof Blob) {
          this.errorMessage = await this.blobErrorMessage(error.response.data, '标准Word校对稿生成失败');
        } else {
          this.errorMessage = error.response?.data?.message || error.message || '标准Word校对稿生成失败';
        }
        this.uploadMessage = '';
      } finally {
        this.isPreparingStandardDocx = false;
      }
    },
    validateBeforeGenerate() {
      if (!this.text || typeof this.text !== 'string' || !this.text.trim()) {
        this.errorMessage = '请先输入或上传正文';
        return false;
      }
      const numericItems = [
        'fontSize',
        'lineSpacing',
        'marginTop',
        'marginBottom',
        'marginLeft',
        'marginRight',
        'lineSpacingSigma',
        'fontSizeSigma',
        'wordSpacingSigma',
        'perturbXSigma',
        'perturbYSigma',
        'perturbThetaSigma',
        'wordSpacing',
        'strikethrough_length_sigma',
        'strikethrough_angle_sigma',
        'strikethrough_width_sigma',
        'strikethrough_probability',
        'strikethrough_width',
        'ink_depth_sigma',
      ];
      for (const item of numericItems) {
        if (Number.isNaN(Number(this[item]))) {
          this.errorMessage = '样式参数必须是数字';
          return false;
        }
      }
      if (this.height < this.marginTop + this.lineSpacing + this.marginBottom && this.isDimensionSpecified) {
        this.errorMessage = '上边距、下边距和行间距之和不能大于高度';
        return false;
      }
      if (this.fontSize > this.lineSpacing) {
        this.errorMessage = '字体大小不能大于行间距';
        return false;
      }
      if (!this.options[this.selectedOption - 1]) {
        this.errorMessage = '请先选择字体';
        return false;
      }
      return true;
    },
    buildGenerationFormData(preview) {
      const formData = new FormData();
      formData.append('text', this.text);
      if (this.options[this.selectedOption - 1]?.text === this.selectedFontFileName && this.fontFile) {
        formData.append('font_file', this.fontFile);
      }
      formData.append('background_image', this.backgroundImage);
      formData.append('font_size', this.fontSize);
      formData.append('line_spacing', this.lineSpacing);
      formData.append('fill', this.fill);
      if (this.width) formData.append('width', this.width);
      if (this.height) formData.append('height', this.height);
      formData.append('top_margin', this.marginTop);
      formData.append('bottom_margin', this.marginBottom);
      formData.append('left_margin', this.marginLeft);
      formData.append('right_margin', this.marginRight);
      formData.append('line_spacing_sigma', this.lineSpacingSigma);
      formData.append('font_size_sigma', this.fontSizeSigma);
      formData.append('word_spacing_sigma', this.wordSpacingSigma);
      formData.append('perturb_x_sigma', this.perturbXSigma);
      formData.append('perturb_y_sigma', this.perturbYSigma);
      formData.append('perturb_theta_sigma', this.perturbThetaSigma);
      formData.append('word_spacing', this.wordSpacing);
      formData.append('preview', preview.toString());
      formData.append('font_option', this.options[this.selectedOption - 1].text);
      formData.append('strikethrough_length_sigma', this.strikethrough_length_sigma);
      formData.append('strikethrough_angle_sigma', this.strikethrough_angle_sigma);
      formData.append('strikethrough_width_sigma', this.strikethrough_width_sigma);
      formData.append('strikethrough_probability', this.strikethrough_probability);
      formData.append('strikethrough_width', this.strikethrough_width);
      formData.append('ink_depth_sigma', this.ink_depth_sigma);
      formData.append('pdf_save', (!preview && this.outputFormat === 'pdf').toString());
      formData.append('output_format', preview ? 'pdf' : this.outputFormat);
      formData.append('content_format', this.activeContentFormat());
      formData.append('isUnderlined', this.isUnderlined.toString());
      formData.append('enableEnglishSpacing', this.enableEnglishSpacing.toString());
      const allowFullPreview = this.isDevEnv && this.enableFullPreview && preview;
      formData.append('full_preview', allowFullPreview.toString());
      return { formData, allowFullPreview };
    },
    async generateHandwriting(preview = false) {
      if (this.isGenerating) {
        this.errorMessage = '正在生成中，请稍候';
        return;
      }
      const currentTime = Date.now();
      const timeSinceLastGenerate = currentTime - this.lastGenerateTime;
      if (timeSinceLastGenerate < this.generateCooldown) {
        this.errorMessage = `请等待 ${Math.ceil((this.generateCooldown - timeSinceLastGenerate) / 1000)} 秒后再次生成`;
        return;
      }
      if (!this.validateBeforeGenerate()) return;

      this.isGenerating = true;
      this.lastGenerateTime = currentTime;
      this.startCooldownTimer();
      this.message = '';
      this.errorMessage = '';
      this.uploadMessage = preview ? '正在生成预览...' : `正在导出 ${this.outputFormatLabel}...`;

      try {
        const { formData, allowFullPreview } = this.buildGenerationFormData(preview);
        const taskCreateResponse = await this.$http.post('/api/generate_handwriting', formData, {
          headers: { 'Content-Type': 'multipart/form-data' },
          withCredentials: true,
        });
        const taskId = taskCreateResponse.data?.task_id;
        if (!taskId) {
          throw new Error('未获取到任务ID');
        }
        this.uploadMessage = `任务已提交，正在生成中（Task ID: ${taskId}）`;
        try {
          await this.waitForTaskViaWebSocket(taskId);
        } catch (wsError) {
          await this.pollGenerationTask(taskId);
        }
        const resultResponse = await this.$http.get(`/api/generate_handwriting/task/${taskId}/result`, {
          responseType: preview ? (allowFullPreview ? 'json' : 'blob') : 'blob',
          withCredentials: true,
        });
        this.handleGenerationResultResponse(resultResponse, preview);
      } catch (error) {
        await this.handleGenerationError(error);
      } finally {
        this.isGenerating = false;
      }
    },
    updateTaskUploadMessage(taskData, taskId) {
      const taskStatus = taskData?.task_status;
      const taskMessage = taskData?.task_message || '任务处理中';
      const taskProgress = taskData?.task_progress;
      const queuePendingCount = taskData?.queue_pending_count;
      const queueAheadCount = taskData?.queue_ahead_count;
      const processingCount = taskData?.processing_count;
      if (taskStatus === 'pending' && typeof queuePendingCount === 'number' && typeof queueAheadCount === 'number') {
        this.uploadMessage = `${taskMessage}（前方 ${queueAheadCount} 人，排队 ${queuePendingCount} 人，处理中 ${processingCount || 0} 人） Task ID: ${taskId}`;
      } else if (typeof taskProgress === 'number') {
        this.uploadMessage = `${taskMessage}（${taskProgress}%） Task ID: ${taskId}`;
      } else {
        this.uploadMessage = `${taskMessage} Task ID: ${taskId}`;
      }
    },
    async waitForTaskViaWebSocket(taskId, timeoutMs = 5 * 60 * 1000) {
      return new Promise((resolve, reject) => {
        let isSettled = false;
        const protocol = window.location.protocol === 'https:' ? 'wss' : 'ws';
        const socket = new WebSocket(`${protocol}://${window.location.host}/api/generate_handwriting/ws/${taskId}`);
        const timeoutId = setTimeout(() => {
          if (isSettled) return;
          isSettled = true;
          socket.close();
          reject(new Error('WebSocket任务等待超时'));
        }, timeoutMs);
        socket.onmessage = (event) => {
          try {
            const data = JSON.parse(event.data);
            if (data?.status === 'error') {
              if (isSettled) return;
              isSettled = true;
              clearTimeout(timeoutId);
              socket.close();
              reject(new Error(data?.message || '任务不存在'));
              return;
            }
            this.updateTaskUploadMessage(data, taskId);
            if (data?.task_status === 'completed') {
              if (isSettled) return;
              isSettled = true;
              clearTimeout(timeoutId);
              socket.close();
              resolve();
            } else if (data?.task_status === 'failed') {
              if (isSettled) return;
              isSettled = true;
              clearTimeout(timeoutId);
              socket.close();
              reject(new Error(data?.error_message || '任务执行失败'));
            }
          } catch (e) {
            // Ignore malformed progress payloads.
          }
        };
        socket.onerror = () => {
          if (isSettled) return;
          isSettled = true;
          clearTimeout(timeoutId);
          reject(new Error('WebSocket连接失败'));
        };
        socket.onclose = () => {
          if (isSettled) return;
          isSettled = true;
          clearTimeout(timeoutId);
          reject(new Error('WebSocket连接已关闭'));
        };
      });
    },
    async pollGenerationTask(taskId, timeoutMs = 5 * 60 * 1000, intervalMs = 1500) {
      const start = Date.now();
      while (Date.now() - start < timeoutMs) {
        const statusResponse = await this.$http.get(`/api/generate_handwriting/task/${taskId}`);
        const taskStatus = statusResponse.data?.task_status;
        this.updateTaskUploadMessage(statusResponse.data, taskId);
        if (taskStatus === 'completed') return;
        if (taskStatus === 'failed') {
          throw new Error(statusResponse.data?.error_message || '任务执行失败');
        }
        await new Promise((resolve) => setTimeout(resolve, intervalMs));
      }
      throw new Error('任务处理超时，请重试');
    },
    filenameFromDisposition(disposition, fallback) {
      if (!disposition) return fallback;
      const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/i);
      if (utf8Match) return decodeURIComponent(utf8Match[1].replace(/"/g, ''));
      const plainMatch = disposition.match(/filename="?([^";]+)"?/i);
      return plainMatch ? plainMatch[1] : fallback;
    },
    downloadBlob(blob, filename) {
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', filename);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
    },
    handleGenerationResultResponse(response, preview) {
      const contentType = response.headers['content-type'] || '';
      if (contentType.includes('application/json') && response.data?.status === 'success') {
        this.previewImages = response.data.images.map((img) => `data:image/png;base64,${img}`);
        this.currentPreviewIndex = 0;
        this.previewImage = this.previewImages[0] || this.previewImage;
        this.message = '预览已生成';
        this.uploadMessage = '';
        return;
      }
      if (contentType.includes('image/png')) {
        const blobUrl = URL.createObjectURL(response.data);
        this.previewImages = [blobUrl];
        this.currentPreviewIndex = 0;
        this.previewImage = blobUrl;
        this.message = '预览已生成';
        this.uploadMessage = '';
        return;
      }
      const suffix = contentType.includes('wordprocessingml') ? 'docx' : 'pdf';
      const fallback = suffix === 'docx' ? 'images.docx' : 'images.pdf';
      const filename = this.filenameFromDisposition(response.headers['content-disposition'], fallback);
      this.downloadBlob(new Blob([response.data], { type: contentType || 'application/octet-stream' }), filename);
      this.message = preview ? '预览已生成' : `文件已导出：${filename}`;
      this.uploadMessage = '';
      this.errorMessage = '';
    },
    async blobErrorMessage(blob, fallback) {
      if (!(blob instanceof Blob)) return fallback;
      try {
        const text = await blob.text();
        const data = JSON.parse(text);
        return data.message || data.error || fallback;
      } catch (error) {
        return fallback;
      }
    },
    async handleGenerationError(error) {
      if (error.response) {
        const errData = error.response.data;
        if (error.response.status === 503 && errData?.status === 'queue_full') {
          this.startQueueFullCountdown(errData.estimated_wait_seconds || 30);
          this.uploadMessage = '';
          return;
        }
        if (errData instanceof Blob) {
          this.errorMessage = await this.blobErrorMessage(errData, '生成失败，请稍后重试');
        } else {
          this.errorMessage = errData?.message || '生成失败，请稍后重试';
        }
      } else {
        this.errorMessage = error.message || '网络错误，请稍后再试';
      }
      this.message = '';
      this.uploadMessage = '';
    },
    startCooldownTimer() {
      this.remainingCooldown = Math.ceil(this.generateCooldown / 1000);
      this.isInCooldownPeriod = true;
      if (this.cooldownTimer) clearInterval(this.cooldownTimer);
      this.cooldownTimer = setInterval(() => {
        this.remainingCooldown -= 1;
        if (this.remainingCooldown <= 0) {
          this.isInCooldownPeriod = false;
          clearInterval(this.cooldownTimer);
          this.cooldownTimer = null;
        }
      }, 1000);
    },
    startQueueFullCountdown(seconds) {
      if (this.queueFullTimer) clearInterval(this.queueFullTimer);
      this.queueFullTotal = seconds;
      this.queueFullCountdown = seconds;
      this.queueFullTimer = setInterval(() => {
        this.queueFullCountdown -= 1;
        if (this.queueFullCountdown <= 0) {
          this.queueFullCountdown = 0;
          clearInterval(this.queueFullTimer);
          this.queueFullTimer = null;
        }
      }, 1000);
    },
    prevPage() {
      if (this.currentPreviewIndex > 0) this.currentPreviewIndex -= 1;
    },
    nextPage() {
      if (this.currentPreviewIndex < this.previewImages.length - 1) this.currentPreviewIndex += 1;
    },
    toggleFullPreview() {
      this.enableFullPreview = !this.enableFullPreview;
    },
    savePreset() {
      const data = {};
      SETTINGS_KEYS.forEach((item) => {
        data[item] = this[item];
        localStorage.setItem(item, JSON.stringify(this[item]));
      });
      localStorage.setItem('myPreset', JSON.stringify(data));
      this.message = '设置已保存';
    },
    loadPreset() {
      const dataString = localStorage.getItem('myPreset');
      if (!dataString || dataString === 'undefined') {
        this.errorMessage = '没有找到保存的设置';
        return;
      }
      try {
        const data = JSON.parse(dataString);
        Object.keys(data).forEach((item) => {
          this[item] = data[item];
        });
        this.message = '设置已载入';
      } catch (error) {
        this.errorMessage = '载入设置失败';
      }
    },
    resetSettings() {
      this.fontFile = null;
      this.backgroundImage = null;
      this.sourceFile = null;
      this.selectedSourceFileName = '';
      this.sourceContentFormat = 'plain';
      this.fontSize = 124;
      this.lineSpacing = 200;
      this.fill = '(0, 0, 0, 255)';
      this.width = 2481;
      this.height = 3507;
      this.marginTop = 50;
      this.marginBottom = 50;
      this.marginLeft = 50;
      this.marginRight = 50;
      this.lineSpacingSigma = 0;
      this.fontSizeSigma = 2;
      this.wordSpacingSigma = 2;
      this.perturbXSigma = 3;
      this.perturbYSigma = 3;
      this.perturbThetaSigma = 0.05;
      this.wordSpacing = 1;
      this.strikethrough_length_sigma = 2;
      this.strikethrough_angle_sigma = 2;
      this.strikethrough_width_sigma = 2;
      this.strikethrough_probability = 0.005;
      this.strikethrough_width = 8;
      this.ink_depth_sigma = 30;
      this.isUnderlined = true;
      this.enableEnglishSpacing = false;
      this.outputFormat = 'pdf';
      this.selectedFontFileName = '';
      this.selectedImageFileName = '';
      this.selectedOption = '1';
      this.previewImage = '/default1.webp';
      this.previewImages = [];
      this.currentPreviewIndex = 0;
      this.message = '设置已重置';
      this.errorMessage = '';
      this.uploadMessage = '';
    },
  },
};
</script>

<style scoped>
.handwriting-workspace {
  min-height: 100vh;
  background:
    linear-gradient(90deg, rgba(34, 77, 86, 0.05) 1px, transparent 1px),
    linear-gradient(180deg, rgba(34, 77, 86, 0.05) 1px, transparent 1px),
    #f7f3ec;
  background-size: 26px 26px;
  color: #203038;
  padding: 24px;
}

.workspace-shell {
  display: grid;
  grid-template-columns: minmax(360px, 560px) minmax(420px, 1fr);
  gap: 24px;
  max-width: 1580px;
  margin: 0 auto;
}

.control-column,
.preview-column {
  min-width: 0;
}

.workspace-header {
  margin-bottom: 18px;
}

.workspace-header p,
.preview-header p {
  margin: 0 0 4px;
  color: #62757c;
  font-size: 0.9rem;
}

.workspace-header h1,
.preview-header h2 {
  margin: 0;
  font-size: clamp(1.45rem, 2.5vw, 2.2rem);
  line-height: 1.18;
  color: #182a31;
}

.panel,
.actions-panel,
.preview-column {
  background: rgba(255, 255, 255, 0.88);
  border: 1px solid rgba(32, 48, 56, 0.12);
  border-radius: 8px;
  box-shadow: 0 14px 32px rgba(31, 46, 54, 0.08);
}

.panel {
  padding: 18px;
  margin-bottom: 16px;
}

.panel-heading,
.preview-header {
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  gap: 16px;
  margin-bottom: 14px;
}

.panel-heading span {
  font-size: 1.08rem;
  font-weight: 700;
}

.panel-heading small {
  color: #6a7c83;
}

.notice {
  border-radius: 8px;
  padding: 10px 12px;
  margin-bottom: 12px;
  font-size: 0.95rem;
}

.notice-success {
  background: #edf8f0;
  color: #1f6635;
  border: 1px solid #bfe1c8;
}

.notice-info {
  background: #eef6fb;
  color: #255978;
  border: 1px solid #bad7e8;
}

.notice-error {
  background: #fff0ef;
  color: #9c2d24;
  border: 1px solid #efc2bd;
}

.source-upload-target {
  position: relative;
  width: 100%;
  min-height: 96px;
}

.source-upload-target.disabled {
  opacity: 0.62;
  cursor: not-allowed;
}

.source-file-input {
  position: absolute;
  inset: 0;
  z-index: 2;
  width: 100%;
  height: 100%;
  opacity: 0;
  cursor: pointer;
}

.source-file-input:disabled {
  cursor: not-allowed;
}

.drop-zone {
  width: 100%;
  min-height: 96px;
  border: 1.5px dashed #7fa0a7;
  border-radius: 8px;
  background: #fbfaf6;
  color: #203038;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 8px;
  pointer-events: none;
  transition: border-color 0.2s ease, transform 0.2s ease, background 0.2s ease;
}

.source-upload-target:hover:not(.disabled) .drop-zone,
.source-upload-target:focus-within:not(.disabled) .drop-zone {
  border-color: #0f6b7a;
  background: #f2faf9;
  transform: translateY(-1px);
}

.drop-zone span {
  color: #6c7c82;
}

.source-meta,
.source-actions,
.asset-row,
.toggle-row,
.actions-panel,
.page-nav {
  display: flex;
  align-items: center;
  gap: 10px;
  flex-wrap: wrap;
}

.source-meta {
  margin-top: 10px;
  color: #496068;
  justify-content: space-between;
}

.source-actions {
  margin-top: 12px;
  justify-content: flex-end;
}

.source-meta button,
.link-button,
.source-actions button {
  border: none;
  background: transparent;
  color: #0f6b7a;
  padding: 0;
}

.source-actions button {
  border: 1px solid #b9c8cb;
  border-radius: 8px;
  background: #ffffff;
  min-height: 40px;
  padding: 8px 14px;
  font-weight: 700;
}

.text-editor {
  width: 100%;
  min-height: 280px;
  margin-top: 14px;
  border: 1px solid #cbd7d9;
  border-radius: 8px;
  padding: 14px;
  resize: vertical;
  line-height: 1.65;
  font-size: 1rem;
  color: #203038;
  background: #fffefa;
}

.settings-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 12px;
}

.settings-grid.compact {
  grid-template-columns: repeat(3, minmax(0, 1fr));
  margin-top: 14px;
}

label {
  color: #3a5058;
  font-weight: 600;
  font-size: 0.92rem;
}

input,
select,
textarea {
  font: inherit;
}

.settings-grid input,
.settings-grid select {
  display: block;
  width: 100%;
  min-height: 42px;
  margin-top: 6px;
  border: 1px solid #cbd7d9;
  border-radius: 8px;
  padding: 8px 10px;
  background: white;
  color: #203038;
}

.asset-row {
  margin-top: 14px;
  color: #62757c;
}

.asset-row button,
.secondary-action,
.primary-action,
.page-nav button {
  border: 1px solid #b9c8cb;
  border-radius: 8px;
  background: #ffffff;
  color: #203038;
  min-height: 40px;
  padding: 8px 14px;
  font-weight: 700;
}

.asset-row button:hover:not(:disabled),
.secondary-action:hover:not(:disabled),
.page-nav button:hover:not(:disabled) {
  border-color: #0f6b7a;
  color: #0f6b7a;
}

button:disabled {
  opacity: 0.55;
  cursor: not-allowed;
}

.toggle-row {
  margin-top: 14px;
  justify-content: space-between;
}

.toggle-row input {
  margin-right: 6px;
}

.advanced-settings {
  margin-top: 14px;
  border-top: 1px solid #e1e7e8;
  padding-top: 12px;
}

.advanced-settings summary {
  cursor: pointer;
  font-weight: 700;
  color: #0f6b7a;
}

.actions-panel {
  padding: 14px;
}

.primary-action {
  background: #0f6b7a;
  border-color: #0f6b7a;
  color: white;
}

.primary-action.export {
  background: #243f49;
  border-color: #243f49;
  flex: 1 1 160px;
}

.secondary-action {
  flex: 0 0 auto;
}

.preview-column {
  padding: 24px;
  position: sticky;
  top: 18px;
  align-self: start;
}

.paper-preview {
  min-height: 70vh;
  border-radius: 8px;
  border: 1px solid #d8e0e2;
  background: #ffffff;
  display: flex;
  align-items: flex-start;
  justify-content: center;
  padding: 20px;
  overflow: auto;
}

.paper-preview img {
  width: min(100%, 760px);
  height: auto;
  border: 1px solid #e2e5e4;
  box-shadow: 0 18px 30px rgba(20, 32, 36, 0.10);
  background: white;
}

.page-nav span {
  color: #4d6067;
  font-weight: 700;
}

.workspace-footer {
  max-width: 1580px;
  margin: 20px auto 0;
  display: flex;
  gap: 14px;
  justify-content: center;
  color: #667980;
}

.workspace-footer a {
  color: #0f6b7a;
  font-weight: 700;
}

@media (max-width: 1080px) {
  .workspace-shell {
    grid-template-columns: 1fr;
  }

  .preview-column {
    position: static;
  }
}

@media (max-width: 680px) {
  .handwriting-workspace {
    padding: 14px;
  }

  .settings-grid,
  .settings-grid.compact {
    grid-template-columns: 1fr;
  }

  .paper-preview {
    min-height: 420px;
    padding: 10px;
  }
}
</style>
