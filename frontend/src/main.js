import "bootstrap/dist/css/bootstrap.css";
import "bootstrap/dist/js/bootstrap.js";
import "bootstrap";
import { createApp } from "vue";
import App from "./App.vue";
import router from "./router";
import store from "./store";
import i18n from "./i18n";
import axios from "axios";
import axiosRetry from "axios-retry";
// eslint-disable-next-line no-unused-vars
import Swal from "sweetalert2";
import { createHead } from "@vueuse/head";

// import Viewer from "v-viewer";H
// import "viewerjs/dist/viewer.css";

const app = createApp(App);
const head = createHead();

// const DEFAULT_TITLE = "作业文档转手写体工作台";

// router.afterEach((to) => {
//   app.nextTick(() => {
//     document.title = to.meta.title || DEFAULT_TITLE;
//   });
// });

app.use(store);
app.use(router);
app.use(i18n);
app.use(head);

// 配置自动重试：网络错误 或 5xx 响应自动重试，最多 3 次，指数退避
axiosRetry(axios, {
  retries: 3,
  retryDelay: axiosRetry.exponentialDelay, // 1s → 2s → 4s
  retryCondition: (error) => {
    // 503 queue_full 是业务状态，不需要重试
    if (
      error.response?.status === 503 &&
      error.response?.data?.status === "queue_full"
    ) {
      return false;
    }
    // 网络错误 或 5xx 服务端错误时重试
    return (
      axiosRetry.isNetworkError(error) ||
      axiosRetry.isRetryableError(error)
    );
  },
  onRetry: (retryCount, error) => {
    console.warn(`请求重试第 ${retryCount} 次，原因：${error.message}`);
  },
});

app.config.globalProperties.$http = axios;
app.config.globalProperties.$swal = Swal;
// app.use(Viewer);

app.mount("#app");

// 注销旧版 /sw.js，迁移到 workbox 生成的 service-worker.js
if ("serviceWorker" in navigator) {
  navigator.serviceWorker.getRegistrations().then((registrations) => {
    for (const registration of registrations) {
      if (registration.active?.scriptURL?.includes("sw.js")) {
        registration.unregister();
      }
    }
  });
}
