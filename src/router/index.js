import Vue from "vue";
import VueRouter from "vue-router";
import Home from "../views/Home.vue";

Vue.use(VueRouter);

const routes = [{
  path: "*",
  name: "Home",
  component: Home,
  meta: {
    title: "隨機抽驗名單工具"
  }
}];

const router = new VueRouter({
  mode: 'history',
  routes,
});

router.beforeEach((to, from, next) => {
  window.document.title = to.meta.title;
  next()
})

export default router;