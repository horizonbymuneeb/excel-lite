/* excel-lite docs — progressive enhancement only. Without JS the site is fully readable:
   navigation, links, tables and code all work; search, theme toggle and copy buttons do not. */
(() => {
  const root = document.documentElement;
  const themeKey = 'excel-lite-theme';

  // ---------------------------------------------------------------- theme
  const stored = (() => {
    try {
      return localStorage.getItem(themeKey);
    } catch {
      return null;
    }
  })();
  if (stored === 'light' || stored === 'dark') root.dataset.theme = stored;
  else root.dataset.theme = window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';

  document.getElementById('theme-toggle')?.addEventListener('click', () => {
    const next = root.dataset.theme === 'dark' ? 'light' : 'dark';
    root.dataset.theme = next;
    try {
      localStorage.setItem(themeKey, next);
    } catch {
      /* private mode */
    }
  });

  // ---------------------------------------------------------------- mobile nav
  const menu = document.querySelector('.menu-button');
  const scrim = document.querySelector('.scrim');
  const setNav = (open) => {
    document.body.classList.toggle('nav-open', open);
    menu?.setAttribute('aria-expanded', String(open));
    if (scrim) scrim.hidden = !open;
  };
  menu?.addEventListener('click', () => setNav(!document.body.classList.contains('nav-open')));
  scrim?.addEventListener('click', () => setNav(false));
  document.querySelectorAll('.sidebar a').forEach((link) => link.addEventListener('click', () => setNav(false)));
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') setNav(false);
  });

  // ---------------------------------------------------------------- copy code
  document.querySelectorAll('.code').forEach((block) => {
    const button = block.querySelector('.code-copy');
    const code = block.querySelector('pre code');
    button?.addEventListener('click', async () => {
      try {
        await navigator.clipboard.writeText(code?.textContent ?? '');
        button.textContent = 'Copied';
        setTimeout(() => (button.textContent = 'Copy'), 1400);
      } catch {
        button.textContent = 'Press ⌘C';
        setTimeout(() => (button.textContent = 'Copy'), 1400);
      }
    });
  });

  // ---------------------------------------------------------------- toc scrollspy
  const tocLinks = [...document.querySelectorAll('.toc a')];
  if (tocLinks.length && 'IntersectionObserver' in window) {
    const byId = new Map(tocLinks.map((link) => [link.getAttribute('href').slice(1), link]));
    const seen = new Set();
    const observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (entry.isIntersecting) seen.add(entry.target.id);
          else seen.delete(entry.target.id);
        }
        const firstVisible = [...document.querySelectorAll('.content h2')].find((heading) => seen.has(heading.id));
        tocLinks.forEach((link) => link.classList.remove('active'));
        if (firstVisible) byId.get(firstVisible.id)?.classList.add('active');
      },
      { rootMargin: '-72px 0px -70% 0px', threshold: 0 },
    );
    document.querySelectorAll('.content h2[id]').forEach((heading) => observer.observe(heading));
  }

  // ---------------------------------------------------------------- search
  const input = document.getElementById('search-input');
  const results = document.getElementById('search-results');
  let indexPromise = null;

  const loadIndex = () => {
    indexPromise ??= fetch('search-index.json')
      .then((response) => response.json())
      .catch(() => []);
    return indexPromise;
  };

  const render = (hits) => {
    if (!results) return;
    results.hidden = false;
    if (hits.length === 0) {
      results.innerHTML = '<p class="empty">No matches. Try “dates”, “freeze”, “memory”, “migrate”.</p>';
      return;
    }
    results.innerHTML = hits
      .map(
        (hit) =>
          `<a href="${hit.url}${hit.id ? `#${hit.id}` : ''}">` +
          `<span class="hit-section">${hit.section}</span>${hit.label}` +
          (hit.context ? `<span class="hit-context">${hit.context}</span>` : '') +
          `</a>`,
      )
      .join('');
  };

  const search = (query) => {
    const term = query.trim().toLowerCase();
    if (term.length < 2) {
      if (results) results.hidden = true;
      return;
    }
    loadIndex().then((pages) => {
      const hits = [];
      for (const page of pages) {
        if (page.title.toLowerCase().includes(term)) {
          hits.push({ url: page.url, label: page.title, section: page.section, context: '' });
        }
        for (const heading of page.headings ?? []) {
          if (heading.text.toLowerCase().includes(term)) {
            hits.push({
              url: page.url,
              id: heading.id,
              label: heading.text,
              section: page.title,
              context: '',
            });
          }
        }
        if (hits.length > 24) break;
      }
      render(hits.slice(0, 12));
    });
  };

  input?.addEventListener('input', (event) => search(event.target.value));
  input?.addEventListener('focus', () => search(input.value));
  document.addEventListener('click', (event) => {
    if (!event.target.closest('.search')) {
      if (results) results.hidden = true;
    }
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === '/' && document.activeElement !== input) {
      event.preventDefault();
      input?.focus();
    }
    if (event.key === 'Escape' && results) results.hidden = true;
  });
})();
