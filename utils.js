// ===================================
// UTILS.JS — Format, Animations, Helpers
// ===================================

const Utils = {
  // Format number as VND currency
  formatVND(amount) {
    return new Intl.NumberFormat('vi-VN', {
      style: 'currency',
      currency: 'VND',
      maximumFractionDigits: 0
    }).format(amount);
  },

  // Format number with dot separator
  formatNumber(num) {
    return new Intl.NumberFormat('vi-VN').format(num);
  },

  // Format date to Vietnamese style
  formatDate(dateStr) {
    const d = new Date(dateStr);
    return d.toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' });
  },

  // Format date short
  formatDateShort(dateStr) {
    const d = new Date(dateStr);
    return d.toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit' });
  },

  // Format time
  formatTime(dateStr) {
    const d = new Date(dateStr);
    return d.toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' });
  },

  // Relative time
  timeAgo(dateStr) {
    const now = new Date();
    const d = new Date(dateStr);
    const diff = Math.floor((now - d) / 1000);
    if (diff < 60) return 'Vừa xong';
    if (diff < 3600) return `${Math.floor(diff / 60)} phút trước`;
    if (diff < 86400) return `${Math.floor(diff / 3600)} giờ trước`;
    if (diff < 604800) return `${Math.floor(diff / 86400)} ngày trước`;
    return Utils.formatDate(dateStr);
  },

  // Animated counter
  animateCounter(element, target, duration = 1200) {
    let start = 0;
    const startTime = performance.now();
    const isVND = element.dataset.format === 'vnd';

    function step(currentTime) {
      const elapsed = currentTime - startTime;
      const progress = Math.min(elapsed / duration, 1);
      // Ease out cubic
      const ease = 1 - Math.pow(1 - progress, 3);
      const current = Math.floor(ease * target);

      if (isVND) {
        element.textContent = Utils.formatVND(current);
      } else {
        element.textContent = Utils.formatNumber(current);
      }

      if (progress < 1) {
        requestAnimationFrame(step);
      }
    }
    requestAnimationFrame(step);
  },

  // Debounce
  debounce(fn, delay = 300) {
    let timer;
    return (...args) => {
      clearTimeout(timer);
      timer = setTimeout(() => fn(...args), delay);
    };
  },

  // Generate unique ID
  uid() {
    return 'id-' + Math.random().toString(36).substr(2, 9);
  },

  // Get today's date string
  today() {
    return new Date().toISOString().split('T')[0];
  },

  // Paginate array
  paginate(array, page, perPage = 10) {
    const start = (page - 1) * perPage;
    return {
      data: array.slice(start, start + perPage),
      totalPages: Math.ceil(array.length / perPage),
      currentPage: page,
      total: array.length
    };
  },

  // Simple search filter
  searchFilter(items, query, fields) {
    if (!query) return items;
    const q = query.toLowerCase();
    return items.filter(item =>
      fields.some(f => String(item[f]).toLowerCase().includes(q))
    );
  },

  // Get week number
  getWeekNumber(d) {
    const date = new Date(d);
    date.setHours(0, 0, 0, 0);
    date.setDate(date.getDate() + 3 - (date.getDay() + 6) % 7);
    const week1 = new Date(date.getFullYear(), 0, 4);
    return 1 + Math.round(((date - week1) / 86400000 - 3 + (week1.getDay() + 6) % 7) / 7);
  },

  // Group data by period
  groupByPeriod(transactions, period = 'month') {
    const groups = {};
    transactions.forEach(t => {
      const d = new Date(t.date);
      let key;
      if (period === 'day') {
        key = d.toISOString().split('T')[0];
      } else if (period === 'week') {
        const wk = Utils.getWeekNumber(d);
        key = `${d.getFullYear()}-T${wk}`;
      } else {
        key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
      }
      if (!groups[key]) groups[key] = [];
      groups[key].push(t);
    });
    return groups;
  },

  // Format period label
  formatPeriodLabel(key, period) {
    if (period === 'day') {
      const d = new Date(key);
      return d.toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit' });
    } else if (period === 'week') {
      return key.replace('-T', ' Tuần ');
    } else {
      const [y, m] = key.split('-');
      const months = ['Th1','Th2','Th3','Th4','Th5','Th6','Th7','Th8','Th9','Th10','Th11','Th12'];
      return months[parseInt(m) - 1] + '/' + y;
    }
  }
};
