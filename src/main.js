import Vue from 'vue';
import App from './App.vue';

Vue.config.productionTip = false;

// Office.js 加载完成后再挂载 Vue
Office.onReady(() => {
	new Vue({
		render: (h) => h(App),
	}).$mount('#app');
});
