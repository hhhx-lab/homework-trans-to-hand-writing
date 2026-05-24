<template>
  <div id="feedBack" class="container mt-5">
    <h1 class="mb-3">本地反馈记录</h1>
    <form @submit.prevent="saveFeedback" class="mt-4">
      <div class="form-group">
        <label for="feedback">问题或改进点</label>
        <textarea
          id="feedback"
          v-model="feedback"
          class="form-control"
          rows="5"
          required
        ></textarea>
      </div>
      <button type="submit" class="btn btn-primary mt-3">保存到本地</button>
    </form>

    <div v-if="savedItems.length" class="mt-4">
      <h2 class="h5">已记录</h2>
      <ul class="list-group">
        <li v-for="item in savedItems" :key="item.createdAt" class="list-group-item">
          <small class="text-muted">{{ item.createdAt }}</small>
          <p class="mb-0">{{ item.content }}</p>
        </li>
      </ul>
    </div>
  </div>
</template>

<script>
const STORAGE_KEY = "handwritingLocalFeedback";

export default {
  name: "UserFeedback",
  data() {
    return {
      feedback: "",
      savedItems: [],
    };
  },
  created() {
    this.loadFeedback();
  },
  methods: {
    loadFeedback() {
      try {
        this.savedItems = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
      } catch (error) {
        this.savedItems = [];
      }
    },
    saveFeedback() {
      const content = this.feedback.trim();
      if (!content) return;
      const item = {
        content,
        createdAt: new Date().toLocaleString("zh-CN"),
      };
      this.savedItems = [item, ...this.savedItems].slice(0, 20);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(this.savedItems));
      this.feedback = "";
    },
  },
};
</script>
