<template>
  <q-layout view="lHh Lpr lFf">
    <q-header elevated>
      <q-toolbar>
        <q-btn
          flat
          dense
          round
          icon="apps"
          aria-label="Menu"
          color="secondary"
          @click="toggleLeftDrawer"
        />

        <q-toolbar-title class="text-bold">MedQuímica FrontEnd</q-toolbar-title>
      </q-toolbar>
    </q-header>

    <q-drawer
      v-model="leftDrawerOpen"
      show-if-above
      class="bg-secondary"
    >
      <q-list>
        <q-item-label class="bg-primary text-white shadow" header>
          Soluções disponiveis !
        </q-item-label>
        <Apps/>
      </q-list>
    </q-drawer>

    <q-page-container>
      <router-view />
    </q-page-container>
  </q-layout>
</template>

<script>
import { defineComponent, ref } from 'vue'
import Apps from 'components/Apps.vue'

export default defineComponent({
  name: 'MainLayout',

  components: {
    Apps
  },

  setup () {
    const leftDrawerOpen = ref(false)

    return {
      leftDrawerOpen,
      toggleLeftDrawer () {
        leftDrawerOpen.value = !leftDrawerOpen.value
      }
    }
  }
})
</script>
