<template>
  <div v-if="showInstallPrompt" class="pwa-install-prompt">
    <div class="prompt-content">
      <div class="prompt-icon">
        <img src="/icon.svg" alt="App Icon" width="48" height="48">
      </div>
      <div class="prompt-text">
        <h3>安装手写生成器</h3>
        <p>将此应用添加到主屏幕，获得更好的使用体验</p>
      </div>
      <div class="prompt-actions">
        <button @click="installApp" class="install-btn">安装</button>
        <button @click="dismissPrompt" class="dismiss-btn">稍后</button>
      </div>
    </div>
  </div>
</template>

<script>
export default {
  name: 'PWAInstallPrompt',
  data() {
    return {
      showInstallPrompt: false,
      deferredPrompt: null
    };
  },
  mounted() {
    const dismissed = this.getCookie('pwa-install-dismissed') === '1';

    // 监听 beforeinstallprompt 事件
    window.addEventListener('beforeinstallprompt', (e) => {
      // 阻止默认的安装提示
      e.preventDefault();

      // 用户明确拒绝后，不再显示
      if (dismissed || this.getCookie('pwa-install-dismissed') === '1') {
        return;
      }

      // 保存事件，稍后使用
      this.deferredPrompt = e;
      window.setTimeout(() => {
        if (this.deferredPrompt && this.getCookie('pwa-install-dismissed') !== '1') {
          this.showInstallPrompt = true;
        }
      }, 120000);
    });

    // 监听应用安装事件
    window.addEventListener('appinstalled', () => {
      this.showInstallPrompt = false;
      this.deferredPrompt = null;
    });
  },
  methods: {
    async installApp() {
      if (!this.deferredPrompt) {
        return;
      }

      // 显示安装提示
      this.deferredPrompt.prompt();
      
      // 等待用户响应 deferredPrompt是之前保存的浏览器事件 测试 测试
      const { outcome } = await this.deferredPrompt.userChoice;

      // 用户在原生提示中拒绝后，记忆选择并不再弹出
      if (outcome === 'dismissed') {
        this.setCookie('pwa-install-dismissed', '1', 3650);
      }
      
      // 清理
      this.deferredPrompt = null;
      this.showInstallPrompt = false;
    },
    dismissPrompt() {
      this.showInstallPrompt = false;
      // 用户选择“稍后”视为拒绝，后续不再弹出
      this.setCookie('pwa-install-dismissed', '1', 3650);
    },
    setCookie(name, value, days) {
      const expires = new Date(Date.now() + days * 24 * 60 * 60 * 1000).toUTCString();
      document.cookie = `${name}=${encodeURIComponent(value)}; expires=${expires}; path=/; SameSite=Lax`;
    },
    getCookie(name) {
      const target = `${name}=`;
      const cookies = document.cookie ? document.cookie.split('; ') : [];
      for (const item of cookies) {
        if (item.startsWith(target)) {
          return decodeURIComponent(item.substring(target.length));
        }
      }
      return null;
    }
  }
}
</script>

<style scoped>
.pwa-install-prompt {
  position: fixed;
  right: 20px;
  bottom: 20px;
  width: min(420px, calc(100vw - 40px));
  background: rgba(248, 252, 255, 0.96);
  border: 1px solid rgba(145, 188, 232, 0.48);
  border-radius: 16px;
  box-shadow: 0 18px 40px rgba(15, 76, 145, 0.18);
  backdrop-filter: blur(12px);
  z-index: 1000;
  animation: slideUp 0.3s ease-out;
  box-sizing: border-box;
}

@keyframes slideUp {
  from {
    transform: translateY(100%);
    opacity: 0;
  }
  to {
    transform: translateY(0);
    opacity: 1;
  }
}

.prompt-content {
  display: flex;
  align-items: center;
  padding: 14px;
  gap: 12px;
}

.prompt-icon img {
  border-radius: 12px;
  filter: drop-shadow(0 8px 16px rgba(11, 99, 206, 0.18));
}

.prompt-text {
  flex: 1;
}

.prompt-text h3 {
  margin: 0 0 4px 0;
  font-size: 16px;
  font-weight: 800;
  color: #102a43;
}

.prompt-text p {
  margin: 0;
  font-size: 14px;
  color: #627d98;
}

.prompt-actions {
  display: flex;
  gap: 8px;
}

.install-btn, .dismiss-btn {
  padding: 8px 16px;
  border: 1px solid #b6d4ef;
  border-radius: 12px;
  font-size: 14px;
  cursor: pointer;
  font-weight: 800;
  transition: background-color 0.2s, border-color 0.2s, color 0.2s;
}

.install-btn {
  background: #0b63ce;
  border-color: #0b63ce;
  color: white;
}

.install-btn:hover {
  background: #083b82;
}

.dismiss-btn {
  background: #ffffff;
  color: #627d98;
}

.dismiss-btn:hover {
  border-color: #0b63ce;
  color: #0b63ce;
  background: #f3f9ff;
}

@media (max-width: 480px) {
  .pwa-install-prompt {
    left: 12px;
    right: 12px;
    bottom: 12px;
    width: auto;
  }
  
  .prompt-content {
    flex-direction: column;
    text-align: center;
  }
  
  .prompt-actions {
    width: 100%;
    justify-content: center;
  }
}
</style>
