<template>
  <div class="home">
    <!-- <img alt="Vue logo" src="../assets/logo.png" /> -->
    <RandomList msg="隨機抽驗名單工具" />
  </div>
</template>

<script>
// @ is an alias to /src
import RandomList from '@/components/RandomList.vue';

export default {
  name: 'Home',
  components: {
    RandomList,
  },
};
</script>
