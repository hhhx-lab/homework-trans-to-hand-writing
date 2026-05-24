<template>
  <transition name="splash-fade">
    <BookSplash v-if="showSplash" @complete="showSplash = false" />
  </transition>
  <!-- <UserLayout v-show="!showSplash"> -->
      <router-view ref="myComponentRef"/>
      <!-- <HomeView /> -->
  <!-- </UserLayout> -->
  <PWAInstallPrompt />
</template>

<script>
// import UserLayout from './views/UserLayout.vue';
import PWAInstallPrompt from './components/PWAInstallPrompt.vue';
import BookSplash from './components/BookSplash.vue';
// import HomeView from './views/HomeView.vue';
import { computed } from 'vue';
import { useRoute } from 'vue-router';
import { useHead } from '@vueuse/head';

export default {
  name: 'App',
  components: {
    // UserLayout,
    PWAInstallPrompt,
    BookSplash,
    // HomeView
  },
  data() {
    return {
      showSplash: !localStorage.getItem('bookSplashShown'),
    };
  },
  setup() {
    const route = useRoute();
    const site = typeof window !== 'undefined' ? window.location.origin : 'http://localhost:8080';
    const defaultTitle = '作业文档转手写体工作台';
    const defaultDesc = '个人私有的作业文档转手写体工具，支持文档抽取、公式整理、手写预览以及 PDF/Word 导出。';

    const title = computed(() => route.meta?.title || defaultTitle);
    const description = computed(() => route.meta?.description || defaultDesc);
    const robots = computed(() => route.meta?.robots || 'noindex, nofollow');
    const canonical = computed(() => site + route.fullPath);

    useHead(() => ({
      title: title.value,
      meta: [
        { name: 'description', content: description.value },
        { name: 'robots', content: robots.value },
        { property: 'og:type', content: 'website' },
        { property: 'og:url', content: canonical.value },
        { property: 'og:title', content: title.value },
        { property: 'og:description', content: description.value },
        { property: 'og:image', content: '/default1.webp' },
        { property: 'twitter:card', content: 'summary_large_image' },
        { property: 'twitter:url', content: canonical.value },
        { property: 'twitter:title', content: title.value },
        { property: 'twitter:description', content: description.value },
        { property: 'twitter:image', content: '/default1.webp' },
      ],
      link: [
        { rel: 'canonical', href: canonical.value },
      ],
    }));

    return {};
  },
};
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
}

nav {
  padding: 30px;
}

nav a {
  font-weight: bold;
  color: #2c3e50;
}

nav a.router-link-exact-active {
  color: #42b983;
}
button {
  transition: all 0.3s ease;
}

button:hover {
  cursor: pointer;
  transform: scale(1.1);
  box-shadow: 0px 0px 8px rgba(0, 0, 0, 0.3);
}

/* 文字不换行，溢出变为省略号 */
.nowrap {
  white-space: nowrap;
  text-overflow: ellipsis;
  overflow: hidden;
}

</style>
