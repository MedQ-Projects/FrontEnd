<template>
  <q-item
    class="bg-white q-ma-sm"
    style="border-radius: 5px"
    v-for="app in appsLink"
    clickable
    tag="a"
    target="_blank"
    @click="goto(app.link)"
  >
    <q-item-section
      v-if="app.icon"
      avatar
    >
      <q-badge :color="app.color" style="font-size: 22px">
        <q-icon :name="app.icon" />
      </q-badge>
    </q-item-section>

    <q-item-section>
      <q-item-label>{{ app.title }}</q-item-label>
      <q-item-label caption>{{ app.caption }}</q-item-label>
    </q-item-section>
  </q-item>
</template>

<script>
import { defineComponent } from 'vue'

export default defineComponent({
  name: 'Apps',
  data(){
    const apps = [
      {title: 'App 1', caption: 'Aplicativo de gerenciamento', icon: 'manage_accounts', link: '/app1', color: 'blue'},
      {title: 'App 2', caption: 'Aplicativo de calculo', icon: 'functions', link: '/app2', color: 'red'},
      {title: 'App 3', caption: 'Aplicativo de RH', icon: 'groups', link: '/app3', color: 'green'}
    ]
    return{
      appsLink: apps
    }
  },
  methods: {
    goto(link){
      this.$router.push(link)
    }
  }
})
</script>
