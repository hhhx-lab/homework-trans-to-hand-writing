import { createRouter, createWebHistory } from "vue-router";

const Home = () => import(/* webpackChunkName: "page-home" */ "../views/HomeView.vue");
const About = () => import(/* webpackChunkName: "page-about" */ "../views/AboutView.vue");
const Login = () => import(/* webpackChunkName: "page-login" */ "../views/UserLogin.vue");
const Register = () => import(/* webpackChunkName: "page-register" */ "../views/UserRegister.vue");
const UserFeedback = () => import(/* webpackChunkName: "page-feedback" */ "@/components/UserFeedback.vue");
const IntroduceComponent = () => import(/* webpackChunkName: "page-introduce" */ "@/components/Introduce.vue");

const routes = [
  {
    path: "/",
    name: "Home",
    component: Home,
    meta: {
      title: "作业文档转手写体工作台",
      description:
        "个人私有的作业文档转手写体工具，支持文档抽取、公式整理、手写预览以及 PDF/Word 导出。",
      robots: "noindex, nofollow",
    },
  },
  {
    path: "/About",
    name: "About",
    component: About,
    meta: {
      title: "关于 - 作业文档转手写体工作台",
      description: "关于个人私有工作台的功能与使用说明。",
      robots: "noindex, nofollow",
    },
  },
  {
    path: "/Login",
    name: "Login",
    component: Login,
    meta: {
      title: "登录 - 作业文档转手写体工作台",
      description: "登录您的账户以使用更多功能。",
      robots: "noindex, nofollow",
    },
  },
  {
    path: "/Register",
    name: "Register",
    component: Register,
    meta: {
      title: "注册 - 作业文档转手写体工作台",
      description: "注册账户，开始使用个人手写体转换工具。",
      robots: "noindex, nofollow",
    },
  },
  {
    path: "/Feedback",
    name: "Feedback",
    component: UserFeedback,
    meta: {
      title: "反馈 - 作业文档转手写体工作台",
      description: "记录个人使用中的问题与改进项。",
      robots: "noindex, nofollow",
    },
  },
  {
    path: "/Introduce",
    name: "Introduce",
    component: IntroduceComponent,
    meta: {
      title: "功能介绍 - 作业文档转手写体工作台",
      description: "了解个人工作台的文档抽取、公式整理和手写导出流程。",
      robots: "noindex, nofollow",
    },
  },
];

const router = createRouter({
  history: createWebHistory(process.env.BASE_URL),
  routes,
});

router.afterEach(() => {
  const routerViewElement = document.querySelector("router-view");
  if (routerViewElement) {
    routerViewElement.scrollIntoView({ behavior: "smooth", block: "start" });
  }
});

export default router;
