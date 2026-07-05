const STORAGE_KEY = 'cocktail-cost-data';
const DATA_VERSION = 1;
const APP_VERSION = '1.0 Demo';

import * as XLSX from 'xlsx';

const state = {
  ingredients: [],
  preparations: [],
  cocktails: [],
  demoSeeded: false,
  editingIngredientId: null,
  editingPreparationId: null,
  editingCocktailId: null,
};

const INGREDIENT_CATEGORIES = [
  'Виски',
  'Джин',
  'Ром',
  'Текила',
  'Водка',
  'Ликер',
  'Вермут',
  'Биттер',
  'Сироп',
  'Кордиал',
  'Сок',
  'Пюре',
  'Кислоты',
  'Газировка',
  'Вино',
  'Пиво',
  'Другое',
];

const DEFAULT_CATEGORY = 'Другое';

const CATEGORY_BADGE_VARIANTS = {
  Виски: 'spirit',
  Джин: 'spirit',
  Ром: 'spirit',
  Текила: 'spirit',
  Водка: 'spirit',
  Ликер: 'liqueur',
  Вермут: 'wine',
  Биттер: 'bitter',
  Сироп: 'sweet',
  Кордиал: 'sweet',
  Сок: 'fresh',
  Пюре: 'fresh',
  Кислоты: 'acid',
  Газировка: 'soda',
  Вино: 'wine',
  Пиво: 'beer',
  Другое: 'default',
};

// ─── Storage ───────────────────────────────────────────────────────────────

function loadData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const data = JSON.parse(raw);
    state.ingredients = (data.ingredients ?? []).map(normalizeIngredient).filter(Boolean);
    state.preparations = (data.preparations ?? []).map(normalizePreparation).filter(Boolean);
    state.cocktails = (data.cocktails ?? []).map(normalizeCocktail).filter(Boolean);
    state.demoSeeded = data.demoSeeded === true;
    backfillIngredientTimestamps();
  } catch {
    state.ingredients = [];
    state.preparations = [];
    state.cocktails = [];
    showToast('Не удалось загрузить данные. Начинаем с пустого списка.', 'error');
  }
}

function normalizeIngredient(ingredient) {
  if (!ingredient || typeof ingredient !== 'object') return null;

  const category = INGREDIENT_CATEGORIES.includes(ingredient.category)
    ? ingredient.category
    : DEFAULT_CATEGORY;

  const bottleVolume = Number(ingredient.bottleVolume);
  const bottlePrice = Number(ingredient.bottlePrice);

  if (!ingredient.id || !ingredient.name || !bottleVolume || bottleVolume <= 0) {
    return null;
  }

  return {
    id: String(ingredient.id),
    name: String(ingredient.name).trim(),
    category,
    bottleVolume,
    bottlePrice: Number.isFinite(bottlePrice) && bottlePrice >= 0 ? bottlePrice : 0,
    createdAt: Number(ingredient.createdAt) || 0,
  };
}

function normalizePreparationItem(item) {
  if (!item?.ingredientId || !item.amountMl) return null;
  const amountMl = Number(item.amountMl);
  if (!Number.isFinite(amountMl) || amountMl <= 0) return null;
  return { ingredientId: String(item.ingredientId), amountMl };
}

function normalizePreparation(preparation) {
  if (!preparation || typeof preparation !== 'object') return null;

  const totalVolume = Number(preparation.totalVolume);
  const items = (preparation.items ?? []).map(normalizePreparationItem).filter(Boolean);

  if (!preparation.id || !preparation.name || !totalVolume || totalVolume <= 0) {
    return null;
  }

  return {
    id: String(preparation.id),
    name: String(preparation.name).trim(),
    totalVolume,
    items,
  };
}

function normalizeCocktailItem(item) {
  if (!item?.amountMl) return null;

  const amountMl = Number(item.amountMl);
  if (!Number.isFinite(amountMl) || amountMl <= 0) return null;

  if (item.preparationId) {
    return { preparationId: String(item.preparationId), amountMl };
  }

  if (item.ingredientId) {
    return { ingredientId: String(item.ingredientId), amountMl };
  }

  return null;
}

function normalizeCocktail(cocktail) {
  if (!cocktail || typeof cocktail !== 'object') return null;

  const items = (cocktail.items ?? []).map(normalizeCocktailItem).filter(Boolean);

  if (!cocktail.id || !cocktail.name) return null;

  return {
    id: String(cocktail.id),
    name: String(cocktail.name).trim(),
    items,
  };
}

function normalizeCategory(value) {
  const trimmed = String(value ?? '').trim();
  if (!trimmed) return DEFAULT_CATEGORY;

  const exact = INGREDIENT_CATEGORIES.find((cat) => cat.toLowerCase() === trimmed.toLowerCase());
  return exact ?? DEFAULT_CATEGORY;
}

function getCategoryBadgeVariant(category) {
  return CATEGORY_BADGE_VARIANTS[category] ?? 'default';
}

function renderCategoryBadge(category) {
  const variant = getCategoryBadgeVariant(category);
  return `<span class="category-badge category-badge--${variant}">${escapeHtml(category)}</span>`;
}

function backfillIngredientTimestamps() {
  const needsBackfill = state.ingredients.some((ingredient) => !ingredient.createdAt);
  if (!needsBackfill) return;

  const baseTime = Date.now();
  state.ingredients.forEach((ingredient, index) => {
    if (!ingredient.createdAt) {
      ingredient.createdAt = baseTime - index;
    }
  });
  saveData();
}

function saveData() {
  try {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        version: DATA_VERSION,
        demoSeeded: state.demoSeeded,
        ingredients: state.ingredients,
        preparations: state.preparations,
        cocktails: state.cocktails,
      })
    );
  } catch {
    showToast('Не удалось сохранить данные. Проверьте свободное место в браузере.', 'error');
  }
}

// ─── Helpers ───────────────────────────────────────────────────────────────

function generateId() {
  if (typeof crypto?.randomUUID === 'function') {
    return crypto.randomUUID();
  }

  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (char) => {
    const random = (Math.random() * 16) | 0;
    const value = char === 'x' ? random : (random & 0x3) | 0x8;
    return value.toString(16);
  });
}

const moneyFormatter = new Intl.NumberFormat('kk-KZ', {
  style: 'currency',
  currency: 'KZT',
  maximumFractionDigits: 0,
});

function formatMoney(value) {
  if (!Number.isFinite(value)) return '—';
  return moneyFormatter.format(value);
}

function formatCostPerMl(price, volume) {
  if (!volume || volume <= 0) return 0;
  return price / volume;
}

function getIngredientById(id) {
  return state.ingredients.find((i) => i.id === id);
}

function getPreparationById(id) {
  return state.preparations.find((p) => p.id === id);
}

function calculatePreparationTotalCost(items) {
  return items.reduce((total, item) => {
    const ingredient = getIngredientById(item.ingredientId);
    if (!ingredient || !item.amountMl) return total;
    const costPerMl = formatCostPerMl(ingredient.bottlePrice, ingredient.bottleVolume);
    return total + costPerMl * item.amountMl;
  }, 0);
}

function calculatePreparationCostPerMl(preparation) {
  const totalCost = calculatePreparationTotalCost(preparation.items);
  return formatCostPerMl(totalCost, preparation.totalVolume);
}

function encodeCocktailItemRef(type, id) {
  return `${type}:${id}`;
}

function parseCocktailItemRef(value) {
  if (!value) return {};
  if (value.startsWith('p:')) {
    return { preparationId: value.slice(2) };
  }
  if (value.startsWith('i:')) {
    return { ingredientId: value.slice(2) };
  }
  return { ingredientId: value };
}

function getCocktailItemRef(item) {
  if (item.preparationId) {
    return encodeCocktailItemRef('p', item.preparationId);
  }
  if (item.ingredientId) {
    return encodeCocktailItemRef('i', item.ingredientId);
  }
  return '';
}

function getCocktailItemCostPerMl(item) {
  if (item.preparationId) {
    const preparation = getPreparationById(item.preparationId);
    if (!preparation) return 0;
    return calculatePreparationCostPerMl(preparation);
  }

  const ingredient = getIngredientById(item.ingredientId);
  if (!ingredient) return 0;
  return formatCostPerMl(ingredient.bottlePrice, ingredient.bottleVolume);
}

function getCocktailItemName(item) {
  if (item.preparationId) {
    return getPreparationById(item.preparationId)?.name ?? 'Удалённая заготовка';
  }
  return getIngredientById(item.ingredientId)?.name ?? 'Удалённый ингредиент';
}

function getCocktailItemCost(item) {
  if (!item?.amountMl) return 0;
  return getCocktailItemCostPerMl(item) * item.amountMl;
}

function isPreparationItem(item) {
  return Boolean(item.preparationId);
}

function findPreparationByName(name) {
  const normalized = name.trim().toLowerCase();
  return state.preparations.find((p) => p.name.toLowerCase() === normalized);
}

function findCocktailByName(name) {
  const normalized = name.trim().toLowerCase();
  return state.cocktails.find((c) => c.name.toLowerCase() === normalized);
}

function ensureDemoIngredient({ key, name, category, bottleVolume, bottlePrice }, ids, createdAt) {
  const existing = findIngredientByName(name);
  if (existing) {
    ids[key] = existing.id;
    return false;
  }

  const ingredient = {
    id: generateId(),
    name,
    category,
    bottleVolume,
    bottlePrice,
    createdAt,
  };
  state.ingredients.push(ingredient);
  ids[key] = ingredient.id;
  return true;
}

const DEMO_INGREDIENTS = [
  { key: 'gin', name: 'Джин Beefeater', category: 'Джин', bottleVolume: 700, bottlePrice: 18500 },
  { key: 'campari', name: 'Campari', category: 'Биттер', bottleVolume: 700, bottlePrice: 11500 },
  {
    key: 'vermouth',
    name: 'Vermouth Cinzano Rosso',
    category: 'Вермут',
    bottleVolume: 1000,
    bottlePrice: 8500,
  },
  {
    key: 'whiskey',
    name: "Jack Daniel's №7",
    category: 'Виски',
    bottleVolume: 700,
    bottlePrice: 16800,
  },
  { key: 'lemon', name: 'Сок лимона', category: 'Сок', bottleVolume: 1000, bottlePrice: 2800 },
  {
    key: 'syrup',
    name: 'Сироп сахарный 1:1',
    category: 'Сироп',
    bottleVolume: 1000,
    bottlePrice: 3200,
  },
  { key: 'egg', name: 'Яичный белок', category: 'Другое', bottleVolume: 1000, bottlePrice: 1800 },
  {
    key: 'tequila',
    name: 'Tequila Olmeca Silver',
    category: 'Текила',
    bottleVolume: 700,
    bottlePrice: 12500,
  },
  { key: 'cointreau', name: 'Cointreau', category: 'Ликер', bottleVolume: 700, bottlePrice: 19000 },
  {
    key: 'raspberry',
    name: 'Малиновое пюре',
    category: 'Пюре',
    bottleVolume: 1000,
    bottlePrice: 4500,
  },
];

const DEMO_PREPARATION = {
  name: 'Berry Bomb Cordial',
  totalVolume: 500,
  items: [
    { ingredientKey: 'raspberry', amountMl: 200 },
    { ingredientKey: 'syrup', amountMl: 120 },
    { ingredientKey: 'lemon', amountMl: 40 },
  ],
};

const DEMO_COCKTAILS = [
  {
    name: 'Negroni',
    items: [
      { ingredientKey: 'gin', amountMl: 30 },
      { ingredientKey: 'campari', amountMl: 30 },
      { ingredientKey: 'vermouth', amountMl: 30 },
    ],
  },
  {
    name: 'Whiskey Sour',
    items: [
      { ingredientKey: 'whiskey', amountMl: 50 },
      { ingredientKey: 'lemon', amountMl: 25 },
      { ingredientKey: 'syrup', amountMl: 20 },
      { ingredientKey: 'egg', amountMl: 25 },
    ],
  },
  {
    name: 'Margarita',
    items: [
      { ingredientKey: 'tequila', amountMl: 50 },
      { ingredientKey: 'cointreau', amountMl: 25 },
      { ingredientKey: 'lemon', amountMl: 25 },
    ],
  },
];

function loadDemoData() {
  if (state.demoSeeded) {
    showToast('Демо-данные уже добавлены', 'error');
    return;
  }

  const ids = {};
  const baseTime = Date.now();
  let addedIngredients = 0;

  DEMO_INGREDIENTS.forEach((item, index) => {
    if (
      ensureDemoIngredient(item, ids, baseTime - index)
    ) {
      addedIngredients += 1;
    }
  });

  if (!findPreparationByName(DEMO_PREPARATION.name)) {
    state.preparations.push({
      id: generateId(),
      name: DEMO_PREPARATION.name,
      totalVolume: DEMO_PREPARATION.totalVolume,
      items: DEMO_PREPARATION.items.map((item) => ({
        ingredientId: ids[item.ingredientKey],
        amountMl: item.amountMl,
      })),
    });
  }

  DEMO_COCKTAILS.forEach((cocktail) => {
    if (findCocktailByName(cocktail.name)) return;

    state.cocktails.push({
      id: generateId(),
      name: cocktail.name,
      items: cocktail.items.map((item) => ({
        ingredientId: ids[item.ingredientKey],
        amountMl: item.amountMl,
      })),
    });
  });

  state.demoSeeded = true;
  saveData();
  renderIngredientsList();
  refreshDependentViews();
  renderHomeDashboard();
  showToast(
    addedIngredients > 0
      ? `Демо-данные добавлены (${addedIngredients} ингредиентов, 1 заготовка, 3 коктейля)`
      : 'Демо-данные добавлены'
  );
}

function findIngredientByName(name) {
  const normalized = name.trim().toLowerCase();
  return state.ingredients.find((i) => i.name.toLowerCase() === normalized);
}

function upsertIngredient(name, volume, price, category) {
  const normalizedCategory =
    category !== undefined ? normalizeCategory(category) : undefined;
  const existing = findIngredientByName(name);
  if (existing) {
    existing.bottleVolume = volume;
    existing.bottlePrice = price;
    if (normalizedCategory !== undefined) {
      existing.category = normalizedCategory;
    }
    return 'updated';
  }

  state.ingredients.push({
    id: generateId(),
    name: name.trim(),
    bottleVolume: volume,
    bottlePrice: price,
    category: normalizedCategory ?? DEFAULT_CATEGORY,
    createdAt: Date.now(),
  });
  return 'added';
}

function calculateCocktailCost(items) {
  return items.reduce((total, item) => {
    if (!item.amountMl) return total;
    return total + getCocktailItemCost(item);
  }, 0);
}

function showToast(message, type = 'success') {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `toast toast--${type}`;
  toast.textContent = message;
  container.appendChild(toast);
  setTimeout(() => toast.remove(), 3000);
}

function refreshDependentViews() {
  if (document.getElementById('home-panel')?.classList.contains('panel--active')) {
    renderHomeDashboard();
  }
  renderPreparationsList();
  renderCocktailsList();
  refreshPreparationIngredientSelects();
  refreshCocktailIngredientSelects();
}

const ICON_PLUS = `<svg class="btn__icon" width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 5v14M5 12h14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>`;

const EMPTY_STATE_ARROW = `<svg class="empty-state__arrow" width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M6 9l6 6 6-6" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`;

const ICON_EXPORT = `<svg class="btn__icon" width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 3v12m0 0l4-4m-4 4l-4-4M5 21h14" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`;

const ICON_IMPORT = `<svg class="btn__icon" width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 21V9m0 0l4 4m-4-4l-4 4M5 7h14" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`;

const ICON_DEMO = `<svg class="btn__icon" width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 3c-4.5 0-8 1.5-8 3.5S7.5 10 12 10s8-1.5 8-3.5S16.5 3 12 3zM4 10.5c0 2 3.6 3.5 8 3.5s8-1.5 8-3.5M4 15c0 2 3.6 3.5 8 3.5s8-1.5 8-3.5M4 19.5V18c0 2 3.6 3.5 8 3.5s8-1.5 8-3.5v1.5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`;

const ICON_EDIT = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M4 20h4l10.5-10.5a2.1 2.1 0 0 0-3-3L5 17v3z" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/></svg>`;

const ICON_DELETE = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M4 7h16M9 7V5h6v2M7 7l1 12h8l1-12" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>`;

const ICON_CLOSE = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>`;

function createCompositionRow({ selectClass, amountClass, optionsHtml, amountMl = '', getCostText, onUpdate }) {
  const row = document.createElement('div');
  row.className = 'cocktail-ingredient-row';
  row.innerHTML = `
    <select class="form__select ${selectClass}">${optionsHtml}</select>
    <input class="form__input ${amountClass}" type="number" min="0.1" step="0.1" placeholder="мл" value="${amountMl}" />
    <button type="button" class="btn btn--danger btn--icon btn-remove-row" title="Удалить" aria-label="Удалить">${ICON_CLOSE}</button>
    <span class="cocktail-ingredient-row__cost"></span>
  `;

  const select = row.querySelector(`.${selectClass}`);
  const amount = row.querySelector(`.${amountClass}`);
  const costLabel = row.querySelector('.cocktail-ingredient-row__cost');

  function updateRowCost() {
    costLabel.textContent = getCostText(select.value, parseFloat(amount.value));
    onUpdate();
  }

  select.addEventListener('change', updateRowCost);
  amount.addEventListener('input', updateRowCost);
  row.querySelector('.btn-remove-row').addEventListener('click', () => {
    row.remove();
    onUpdate();
  });

  updateRowCost();
  return row;
}

// ─── Tabs ──────────────────────────────────────────────────────────────────

function initTabs() {
  const buttons = document.querySelectorAll('.tabs__btn');
  const panels = {
    home: document.getElementById('home-panel'),
    ingredients: document.getElementById('ingredients-panel'),
    preparations: document.getElementById('preparations-panel'),
    cocktails: document.getElementById('cocktails-panel'),
  };

  buttons.forEach((btn) => {
    btn.addEventListener('click', () => {
      const tab = btn.dataset.tab;

      buttons.forEach((b) => {
        b.classList.toggle('tabs__btn--active', b === btn);
        b.setAttribute('aria-selected', b === btn ? 'true' : 'false');
      });

      Object.entries(panels).forEach(([key, panel]) => {
        if (!panel) return;
        const active = key === tab;
        panel.classList.toggle('panel--active', active);
        panel.hidden = !active;
      });

      if (tab === 'home') {
        renderHomeDashboard();
      }
    });
  });
}

// ─── Home dashboard ────────────────────────────────────────────────────────

function getAverageCocktailCost() {
  if (state.cocktails.length === 0) return null;
  const total = state.cocktails.reduce(
    (sum, cocktail) => sum + calculateCocktailCost(cocktail.items),
    0
  );
  return total / state.cocktails.length;
}

function getMostExpensiveCocktail() {
  if (state.cocktails.length === 0) return null;

  return state.cocktails.reduce((mostExpensive, cocktail) => {
    const cost = calculateCocktailCost(cocktail.items);
    const maxCost = calculateCocktailCost(mostExpensive.items);
    return cost > maxCost ? cocktail : mostExpensive;
  });
}

function getRecentIngredients(limit = 5) {
  return [...state.ingredients]
    .sort((a, b) => {
      const timeDiff = (b.createdAt || 0) - (a.createdAt || 0);
      if (timeDiff !== 0) return timeDiff;
      return state.ingredients.indexOf(b) - state.ingredients.indexOf(a);
    })
    .slice(0, limit);
}

function renderHomeDashboard() {
  const container = document.getElementById('home-dashboard');
  if (!container) return;

  const averageCost = getAverageCocktailCost();
  const mostExpensive = getMostExpensiveCocktail();
  const mostExpensiveCost = mostExpensive ? calculateCocktailCost(mostExpensive.items) : null;
  const recentIngredients = getRecentIngredients();

  const recentHtml =
    recentIngredients.length === 0
      ? '<p class="dashboard__empty">Пока нет ингредиентов</p>'
      : recentIngredients
          .map((ingredient) => {
            const costPerMl = formatCostPerMl(ingredient.bottlePrice, ingredient.bottleVolume);
            return `
              <div class="dashboard-recent-item">
                <div class="dashboard-recent-item__info">
                  <p class="dashboard-recent-item__name">${escapeHtml(ingredient.name)}</p>
                  <p class="dashboard-recent-item__meta">
                    ${ingredient.bottleVolume} мл · ${formatMoney(ingredient.bottlePrice)}
                  </p>
                </div>
                <div class="dashboard-recent-item__aside">
                  ${renderCategoryBadge(ingredient.category)}
                  <span class="dashboard-recent-item__cost">${formatMoney(costPerMl)}/мл</span>
                </div>
              </div>
            `;
          })
          .join('');

  container.innerHTML = `
    <div class="dashboard__hero">
      <div>
        <h2 class="dashboard__title">Обзор</h2>
        <p class="dashboard__subtitle">Сводка по ингредиентам, заготовкам и коктейлям</p>
      </div>
    </div>

    <div class="dashboard__stats">
      <article class="stat-card">
        <span class="stat-card__label">Ингредиенты</span>
        <span class="stat-card__value">${state.ingredients.length}</span>
        <span class="stat-card__hint">в базе</span>
      </article>
      <article class="stat-card">
        <span class="stat-card__label">Заготовки</span>
        <span class="stat-card__value">${state.preparations.length}</span>
        <span class="stat-card__hint">рецептов</span>
      </article>
      <article class="stat-card">
        <span class="stat-card__label">Коктейли</span>
        <span class="stat-card__value">${state.cocktails.length}</span>
        <span class="stat-card__hint">в меню</span>
      </article>
    </div>

    <div class="dashboard__insights">
      <article class="insight-card">
        <span class="insight-card__label">Средняя себестоимость</span>
        <span class="insight-card__value">${averageCost !== null ? formatMoney(averageCost) : '—'}</span>
        <span class="insight-card__hint">на коктейль</span>
      </article>
      <article class="insight-card insight-card--accent">
        <span class="insight-card__label">Самый дорогой коктейль</span>
        <span class="insight-card__value">${mostExpensiveCost !== null ? formatMoney(mostExpensiveCost) : '—'}</span>
        <span class="insight-card__hint">${mostExpensive ? escapeHtml(mostExpensive.name) : 'нет данных'}</span>
      </article>
    </div>

    <article class="card dashboard__recent">
      <div class="card__header">
        <h2 class="card__title">Последние ингредиенты</h2>
        <span class="badge">${recentIngredients.length}</span>
      </div>
      <div class="dashboard__recent-list">${recentHtml}</div>
    </article>

    <article class="card dashboard__demo">
      <div class="card__header">
        <h2 class="card__title">Быстрый старт</h2>
      </div>
      <p class="dashboard__demo-text">
        10 ингредиентов, заготовка Berry Bomb Cordial и коктейли Negroni, Whiskey Sour, Margarita с реалистичными ценами в ₸.
      </p>
      <button
        type="button"
        class="btn btn--secondary"
        id="load-demo-data-btn"
        ${state.demoSeeded ? 'disabled' : ''}
      >
        ${ICON_DEMO}
        ${state.demoSeeded ? 'Демо-данные уже добавлены' : 'Заполнить демо-данными'}
      </button>
    </article>

    <article class="card dashboard__backup">
      <div class="card__header">
        <h2 class="card__title">Резервная копия</h2>
      </div>
      <p class="dashboard__backup-text">
        Сохраните все данные в JSON-файл, чтобы перенести их на другой компьютер или восстановить после сбоя.
      </p>
      <div class="dashboard__backup-actions">
        <button type="button" class="btn btn--secondary" id="backup-export-btn">
          ${ICON_EXPORT}
          Экспорт резервной копии
        </button>
        <button type="button" class="btn btn--secondary" id="backup-import-btn">
          ${ICON_IMPORT}
          Импорт резервной копии
        </button>
      </div>
    </article>
  `;

  bindBackupControls();
  bindDemoControls();
}

const BACKUP_VERSION = 1;

function createBackupPayload() {
  return {
    version: BACKUP_VERSION,
    app: 'zhaslab',
    appVersion: APP_VERSION,
    demoSeeded: state.demoSeeded,
    exportedAt: new Date().toISOString(),
    ingredients: state.ingredients,
    preparations: state.preparations,
    cocktails: state.cocktails,
  };
}

function exportBackup() {
  const payload = createBackupPayload();
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  const date = new Date().toISOString().slice(0, 10);

  link.href = url;
  link.download = `zhaslab-backup-${date}.json`;
  link.click();
  URL.revokeObjectURL(url);
  showToast('Резервная копия сохранена');
}

function parseBackupFile(data) {
  if (!data || typeof data !== 'object') {
    throw new Error('Неверный формат файла резервной копии');
  }

  const ingredients = Array.isArray(data.ingredients) ? data.ingredients : [];
  const preparations = Array.isArray(data.preparations) ? data.preparations : [];
  const cocktails = Array.isArray(data.cocktails) ? data.cocktails : [];

  if (ingredients.length === 0 && preparations.length === 0 && cocktails.length === 0) {
    throw new Error('Файл резервной копии пуст');
  }

  const normalized = {
    ingredients: ingredients.map(normalizeIngredient).filter(Boolean),
    preparations: preparations.map(normalizePreparation).filter(Boolean),
    cocktails: cocktails.map(normalizeCocktail).filter(Boolean),
  };

  if (
    normalized.ingredients.length === 0 &&
    normalized.preparations.length === 0 &&
    normalized.cocktails.length === 0
  ) {
    throw new Error('Файл не содержит корректных данных');
  }

  return normalized;
}

async function handleBackupImport(file) {
  const text = await file.text();
  const raw = JSON.parse(text);
  const backup = parseBackupFile(raw);

  const message =
    'Импорт заменит все текущие данные:\n' +
    `• ${backup.ingredients.length} ингредиентов\n` +
    `• ${backup.preparations.length} заготовок\n` +
    `• ${backup.cocktails.length} коктейлей\n\n` +
    'Продолжить?';

  if (!confirm(message)) return;

  state.ingredients = backup.ingredients;
  state.preparations = backup.preparations;
  state.cocktails = backup.cocktails;
  state.demoSeeded = raw.demoSeeded === true;
  state.editingIngredientId = null;
  state.editingPreparationId = null;
  state.editingCocktailId = null;

  saveData();
  resetIngredientForm();
  resetPreparationForm();
  resetCocktailForm();
  populateCategorySelects();
  renderIngredientsList();
  refreshDependentViews();
  showToast('Резервная копия успешно восстановлена');
}

function bindBackupControls() {
  const exportBtn = document.getElementById('backup-export-btn');
  const importBtn = document.getElementById('backup-import-btn');

  if (exportBtn && !exportBtn.dataset.bound) {
    exportBtn.dataset.bound = 'true';
    exportBtn.addEventListener('click', exportBackup);
  }

  if (importBtn && !importBtn.dataset.bound) {
    importBtn.dataset.bound = 'true';
    importBtn.addEventListener('click', () => {
      document.getElementById('backup-import-input')?.click();
    });
  }
}

function bindDemoControls() {
  const btn = document.getElementById('load-demo-data-btn');
  if (!btn || btn.dataset.bound || btn.disabled) return;

  btn.dataset.bound = 'true';
  btn.addEventListener('click', loadDemoData);
}

function initBackupImport() {
  const input = document.getElementById('backup-import-input');
  if (!input || input.dataset.bound) return;

  input.dataset.bound = 'true';
  input.addEventListener('change', async () => {
    const file = input.files?.[0];
    input.value = '';
    if (!file) return;

    const isJson =
      file.name.toLowerCase().endsWith('.json') || file.type === 'application/json';

    if (!isJson) {
      showToast('Выберите JSON-файл резервной копии', 'error');
      return;
    }

    try {
      await handleBackupImport(file);
    } catch (error) {
      showToast(error.message || 'Не удалось импортировать резервную копию', 'error');
    }
  });
}

// ─── Ingredients ───────────────────────────────────────────────────────────

function updateIngredientPreview() {
  const volume = parseFloat(document.getElementById('ingredient-volume').value);
  const price = parseFloat(document.getElementById('ingredient-price').value);
  const preview = document.getElementById('ingredient-preview');
  const valueEl = document.getElementById('ingredient-cost-per-ml');

  if (volume > 0 && price >= 0 && !isNaN(volume) && !isNaN(price)) {
    preview.hidden = false;
    valueEl.textContent = formatMoney(formatCostPerMl(price, volume));
  } else {
    preview.hidden = true;
  }
}

function resetIngredientForm() {
  state.editingIngredientId = null;
  document.getElementById('ingredient-id').value = '';
  document.getElementById('ingredient-name').value = '';
  document.getElementById('ingredient-category').value = DEFAULT_CATEGORY;
  document.getElementById('ingredient-volume').value = '';
  document.getElementById('ingredient-price').value = '';
  document.getElementById('ingredient-preview').hidden = true;
  document.getElementById('ingredient-form-title').textContent = 'Добавить ингредиент';
  document.getElementById('ingredient-submit').textContent = 'Добавить';
  document.getElementById('ingredient-cancel').hidden = true;
}

function getIngredientFilters() {
  const searchInput = document.getElementById('ingredient-search');
  const categoryFilter = document.getElementById('ingredient-category-filter');

  return {
    query: searchInput?.value.trim().toLowerCase() ?? '',
    category: categoryFilter?.value ?? '',
  };
}

function getFilteredIngredients() {
  const { query, category } = getIngredientFilters();

  return state.ingredients.filter((ing) => {
    const matchesSearch = !query || ing.name.toLowerCase().includes(query);
    const matchesCategory = !category || ing.category === category;
    return matchesSearch && matchesCategory;
  });
}

function getEmptyIngredientsMessage(query, category) {
  if (query && category) {
    return `Ничего не найдено по запросу «${escapeHtml(query)}» в категории «${escapeHtml(category)}»`;
  }
  if (query) {
    return `Ничего не найдено по запросу «${escapeHtml(query)}»`;
  }
  if (category) {
    return `Нет ингредиентов в категории «${escapeHtml(category)}»`;
  }
  return 'Пока нет ингредиентов. Добавьте первый!';
}

function renderIngredientsEmptyState() {
  return `
    <div class="empty-state">
      <p class="empty-state__title">Пока нет ингредиентов</p>
      ${EMPTY_STATE_ARROW}
      <p class="empty-state__hint">Добавьте первый ингредиент</p>
      ${EMPTY_STATE_ARROW}
      <button type="button" class="btn btn--primary empty-state__cta" id="empty-add-ingredient-btn">
        ${ICON_PLUS}
        Добавить ингредиент
      </button>
    </div>
  `;
}

function focusIngredientForm() {
  resetIngredientForm();
  document.querySelector('[data-tab="ingredients"]')?.click();
  const formCard = document.getElementById('ingredient-form')?.closest('.card');
  formCard?.scrollIntoView({ behavior: 'smooth', block: 'center' });
  document.getElementById('ingredient-name')?.focus({ preventScroll: true });
}

function bindIngredientsEmptyState() {
  const btn = document.getElementById('empty-add-ingredient-btn');
  if (!btn || btn.dataset.bound) return;
  btn.dataset.bound = 'true';
  btn.addEventListener('click', focusIngredientForm);
}

function populateCategorySelects() {
  const formSelect = document.getElementById('ingredient-category');
  const filterSelect = document.getElementById('ingredient-category-filter');

  formSelect.innerHTML = INGREDIENT_CATEGORIES.map(
    (category) => `<option value="${category}">${category}</option>`
  ).join('');

  filterSelect.innerHTML =
    '<option value="">Все категории</option>' +
    INGREDIENT_CATEGORIES.map(
      (category) => `<option value="${category}">${category}</option>`
    ).join('');
}

function renderIngredientsList() {
  const list = document.getElementById('ingredients-list');
  const count = document.getElementById('ingredients-count');
  const searchClear = document.getElementById('ingredient-search-clear');
  const { query, category } = getIngredientFilters();

  count.textContent = state.ingredients.length;

  const filters = document.getElementById('ingredients-filters');

  if (searchClear) {
    searchClear.hidden = !query;
  }

  if (filters) {
    filters.hidden = state.ingredients.length === 0;
  }

  if (state.ingredients.length === 0) {
    list.innerHTML = renderIngredientsEmptyState();
    bindIngredientsEmptyState();
    return;
  }

  const filtered = getFilteredIngredients();

  if (filtered.length === 0) {
    list.innerHTML = `<p class="list__empty">${getEmptyIngredientsMessage(query, category)}</p>`;
    return;
  }

  list.innerHTML = filtered
    .map((ing) => {
      const costPerMl = formatCostPerMl(ing.bottlePrice, ing.bottleVolume);
      return `
        <div class="list-item" data-id="${ing.id}">
          <div class="list-item__info">
            <div class="list-item__title-row">
              <p class="list-item__name">${escapeHtml(ing.name)}</p>
              ${renderCategoryBadge(ing.category)}
            </div>
            <p class="list-item__meta">
              ${ing.bottleVolume} мл · ${formatMoney(ing.bottlePrice)} ·
              ${formatMoney(costPerMl)}/мл
            </p>
          </div>
          <div class="list-item__actions">
            <button class="btn btn--ghost btn--icon btn-edit-ingredient" data-id="${ing.id}" title="Редактировать" aria-label="Редактировать">
              ${ICON_EDIT}
            </button>
            <button class="btn btn--danger btn--icon btn-delete-ingredient" data-id="${ing.id}" title="Удалить" aria-label="Удалить">
              ${ICON_DELETE}
            </button>
          </div>
        </div>
      `;
    })
    .join('');

  list.querySelectorAll('.btn-edit-ingredient').forEach((btn) => {
    btn.addEventListener('click', () => editIngredient(btn.dataset.id));
  });

  list.querySelectorAll('.btn-delete-ingredient').forEach((btn) => {
    btn.addEventListener('click', () => deleteIngredient(btn.dataset.id));
  });
}

function editIngredient(id) {
  const ing = getIngredientById(id);
  if (!ing) return;

  state.editingIngredientId = id;
  document.getElementById('ingredient-id').value = id;
  document.getElementById('ingredient-name').value = ing.name;
  document.getElementById('ingredient-category').value = ing.category ?? DEFAULT_CATEGORY;
  document.getElementById('ingredient-volume').value = ing.bottleVolume;
  document.getElementById('ingredient-price').value = ing.bottlePrice;
  document.getElementById('ingredient-form-title').textContent = 'Редактировать ингредиент';
  document.getElementById('ingredient-submit').textContent = 'Сохранить';
  document.getElementById('ingredient-cancel').hidden = false;
  updateIngredientPreview();
  document.querySelector('[data-tab="ingredients"]')?.click();
  document.getElementById('ingredient-name').focus();
}

function deleteIngredient(id) {
  const ing = getIngredientById(id);
  if (!ing) return;

  const usedInCocktails = state.cocktails.filter((c) =>
    c.items.some((item) => item.ingredientId === id)
  );
  const usedInPreparations = state.preparations.filter((p) =>
    p.items.some((item) => item.ingredientId === id)
  );

  let message = `Удалить «${ing.name}»?`;
  if (usedInCocktails.length > 0) {
    message += `\n\nИнгредиент используется в ${usedInCocktails.length} ${pluralize(usedInCocktails.length, ['коктейле', 'коктейлях', 'коктейлях'])}. Он будет удалён из рецептов.`;
  }
  if (usedInPreparations.length > 0) {
    message += `\n\nИнгредиент используется в ${usedInPreparations.length} ${pluralize(usedInPreparations.length, ['заготовке', 'заготовках', 'заготовках'])}. Он будет удалён из состава.`;
  }

  if (!confirm(message)) return;

  state.ingredients = state.ingredients.filter((i) => i.id !== id);
  state.preparations = state.preparations.map((p) => ({
    ...p,
    items: p.items.filter((item) => item.ingredientId !== id),
  }));
  state.cocktails = state.cocktails.map((c) => ({
    ...c,
    items: c.items.filter((item) => item.ingredientId !== id),
  }));

  if (state.editingIngredientId === id) resetIngredientForm();
  saveData();
  renderIngredientsList();
  refreshDependentViews();
  showToast('Ингредиент удалён');
}

function handleIngredientSubmit(e) {
  e.preventDefault();

  const name = document.getElementById('ingredient-name').value.trim();
  const category = normalizeCategory(document.getElementById('ingredient-category').value);
  const volume = parseFloat(document.getElementById('ingredient-volume').value);
  const price = parseFloat(document.getElementById('ingredient-price').value);

  if (!name) {
    showToast('Введите название ингредиента', 'error');
    return;
  }
  if (!volume || volume <= 0) {
    showToast('Объём бутылки должен быть больше 0', 'error');
    return;
  }
  if (isNaN(price) || price < 0) {
    showToast('Введите корректную цену', 'error');
    return;
  }

  const duplicate = state.ingredients.find(
    (i) => i.name.toLowerCase() === name.toLowerCase() && i.id !== state.editingIngredientId
  );
  if (duplicate) {
    showToast('Ингредиент с таким названием уже существует', 'error');
    return;
  }

  if (state.editingIngredientId) {
    const index = state.ingredients.findIndex((i) => i.id === state.editingIngredientId);
    if (index !== -1) {
      state.ingredients[index] = {
        ...state.ingredients[index],
        name,
        category,
        bottleVolume: volume,
        bottlePrice: price,
      };
    }
    showToast('Ингредиент обновлён');
  } else {
    state.ingredients.push({
      id: generateId(),
      name,
      category,
      bottleVolume: volume,
      bottlePrice: price,
      createdAt: Date.now(),
    });
    showToast('Ингредиент добавлен');
  }

  saveData();
  resetIngredientForm();
  renderIngredientsList();
  refreshDependentViews();
}

function initIngredientForm() {
  const form = document.getElementById('ingredient-form');
  const volumeInput = document.getElementById('ingredient-volume');
  const priceInput = document.getElementById('ingredient-price');

  form.addEventListener('submit', handleIngredientSubmit);
  volumeInput.addEventListener('input', updateIngredientPreview);
  priceInput.addEventListener('input', updateIngredientPreview);
  document.getElementById('ingredient-cancel').addEventListener('click', resetIngredientForm);
  populateCategorySelects();
  initIngredientImport();
  initIngredientSearch();
  initIngredientFilter();
}

function initIngredientFilter() {
  const categoryFilter = document.getElementById('ingredient-category-filter');
  categoryFilter.addEventListener('change', renderIngredientsList);
}

function debounce(fn, delay = 200) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), delay);
  };
}

function initIngredientSearch() {
  const searchInput = document.getElementById('ingredient-search');
  const searchClear = document.getElementById('ingredient-search-clear');

  searchInput.addEventListener('input', debounce(renderIngredientsList));

  searchClear.addEventListener('click', () => {
    searchInput.value = '';
    searchInput.focus();
    renderIngredientsList();
  });
}

// ─── Ingredient import ───────────────────────────────────────────────────────

const IMPORT_COLUMNS = {
  name: ['название'],
  volume: ['объем (мл)', 'объём (мл)', 'объем', 'объём'],
  price: ['цена (₸)', 'цена'],
  category: ['категория', 'category'],
};

function normalizeHeader(header) {
  return String(header).trim().toLowerCase().replace(/\s+/g, ' ');
}

function getImportValue(row, aliases) {
  const normalizedRow = {};
  for (const [key, value] of Object.entries(row)) {
    normalizedRow[normalizeHeader(key)] = value;
  }

  for (const alias of aliases) {
    const value = normalizedRow[alias];
    if (value !== '' && value !== null && value !== undefined) {
      return value;
    }
  }

  for (const alias of aliases) {
    if (alias in normalizedRow) {
      return normalizedRow[alias];
    }
  }

  return undefined;
}

function parseImportNumber(value) {
  if (typeof value === 'number') return value;
  const cleaned = String(value).replace(/\s/g, '').replace(',', '.');
  return parseFloat(cleaned);
}

function parseIngredientRows(rows) {
  const stats = { added: 0, updated: 0, skipped: 0 };

  rows.forEach((row) => {
    const name = String(getImportValue(row, IMPORT_COLUMNS.name) ?? '').trim();
    const volume = parseImportNumber(getImportValue(row, IMPORT_COLUMNS.volume));
    const price = parseImportNumber(getImportValue(row, IMPORT_COLUMNS.price));
    const categoryRaw = getImportValue(row, IMPORT_COLUMNS.category);
    const hasCategory = categoryRaw !== undefined && String(categoryRaw).trim() !== '';

    if (!name && (isNaN(volume) || volume <= 0) && (isNaN(price) || price < 0)) {
      return;
    }

    if (!name) {
      stats.skipped += 1;
      return;
    }
    if (!volume || volume <= 0 || isNaN(volume)) {
      stats.skipped += 1;
      return;
    }
    if (isNaN(price) || price < 0) {
      stats.skipped += 1;
      return;
    }

    const result = upsertIngredient(
      name,
      volume,
      price,
      hasCategory ? categoryRaw : undefined
    );
    stats[result] += 1;
  });

  if (stats.added === 0 && stats.updated === 0 && stats.skipped === rows.length) {
    throw new Error('Не найдено подходящих строк. Проверьте заголовки столбцов в Excel.');
  }

  return stats;
}

async function handleIngredientImport(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error('Файл Excel пуст');
  }

  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  if (rows.length === 0) {
    throw new Error('В файле нет данных для импорта');
  }

  return parseIngredientRows(rows);
}

function initIngredientImport() {
  const importBtn = document.getElementById('ingredient-import');
  const importInput = document.getElementById('ingredient-import-input');

  importBtn.addEventListener('click', () => importInput.click());

  importInput.addEventListener('change', async () => {
    const file = importInput.files?.[0];
    importInput.value = '';

    if (!file) return;

    const isXlsx =
      file.name.toLowerCase().endsWith('.xlsx') ||
      file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    if (!isXlsx) {
      showToast('Выберите файл Excel (.xlsx)', 'error');
      return;
    }

    importBtn.disabled = true;
    importBtn.classList.add('btn--loading');

    try {
      const stats = await handleIngredientImport(file);
      saveData();
      renderIngredientsList();
      refreshDependentViews();

      const parts = [];
      if (stats.added > 0) parts.push(`добавлено: ${stats.added}`);
      if (stats.updated > 0) parts.push(`обновлено: ${stats.updated}`);
      if (stats.skipped > 0) parts.push(`пропущено: ${stats.skipped}`);

      showToast(`Импорт завершён (${parts.join(', ')})`);
    } catch (error) {
      showToast(error.message || 'Не удалось импортировать файл', 'error');
    } finally {
      importBtn.disabled = false;
      importBtn.classList.remove('btn--loading');
    }
  });
}

// ─── Preparations ────────────────────────────────────────────────────────────

function buildBaseIngredientOptions(selectedId = '') {
  if (state.ingredients.length === 0) {
    return '<option value="">— Сначала добавьте ингредиенты —</option>';
  }

  return (
    '<option value="">— Выберите ингредиент —</option>' +
    state.ingredients
      .map(
        (ing) =>
          `<option value="${ing.id}" ${ing.id === selectedId ? 'selected' : ''}>${escapeHtml(ing.name)}</option>`
      )
      .join('')
  );
}

function createPreparationIngredientRow(ingredientId = '', amountMl = '') {
  return createCompositionRow({
    selectClass: 'preparation-ingredient-select',
    amountClass: 'preparation-ingredient-amount',
    optionsHtml: buildBaseIngredientOptions(ingredientId),
    amountMl,
    getCostText: (ingredientIdValue, ml) => {
      const ingredient = getIngredientById(ingredientIdValue);
      if (ingredient && ml > 0) {
        return `= ${formatMoney(formatCostPerMl(ingredient.bottlePrice, ingredient.bottleVolume) * ml)}`;
      }
      return '';
    },
    onUpdate: updatePreparationPreview,
  });
}

function refreshPreparationIngredientSelects() {
  const container = document.getElementById('preparation-ingredients');
  if (!container) return;

  container.querySelectorAll('.cocktail-ingredient-row').forEach((row) => {
    const select = row.querySelector('.preparation-ingredient-select');
    const current = select.value;
    select.innerHTML = buildBaseIngredientOptions(current);
  });
}

function getPreparationFormItems() {
  const rows = document.querySelectorAll('#preparation-ingredients .cocktail-ingredient-row');
  return Array.from(rows)
    .map((row) => ({
      ingredientId: row.querySelector('.preparation-ingredient-select').value,
      amountMl: parseFloat(row.querySelector('.preparation-ingredient-amount').value),
    }))
    .filter((item) => item.ingredientId && item.amountMl > 0);
}

function updatePreparationPreview() {
  const items = getPreparationFormItems();
  const volume = parseFloat(document.getElementById('preparation-volume').value);
  const costPreview = document.getElementById('preparation-cost-preview');
  const mlPreview = document.getElementById('preparation-ml-preview');
  const totalEl = document.getElementById('preparation-total-cost');
  const perMlEl = document.getElementById('preparation-cost-per-ml');
  const totalCost = calculatePreparationTotalCost(items);

  if (items.length > 0) {
    costPreview.hidden = false;
    totalEl.textContent = formatMoney(totalCost);
  } else {
    costPreview.hidden = true;
  }

  if (items.length > 0 && volume > 0 && !isNaN(volume)) {
    mlPreview.hidden = false;
    perMlEl.textContent = formatMoney(formatCostPerMl(totalCost, volume));
  } else {
    mlPreview.hidden = true;
  }
}

function resetPreparationForm() {
  state.editingPreparationId = null;
  document.getElementById('preparation-id').value = '';
  document.getElementById('preparation-name').value = '';
  document.getElementById('preparation-volume').value = '';
  document.getElementById('preparation-ingredients').innerHTML = '';
  document.getElementById('preparation-cost-preview').hidden = true;
  document.getElementById('preparation-ml-preview').hidden = true;
  document.getElementById('preparation-form-title').textContent = 'Создать заготовку';
  document.getElementById('preparation-submit').textContent = 'Создать';
  document.getElementById('preparation-cancel').hidden = true;
  addPreparationIngredientRow();
}

function addPreparationIngredientRow(ingredientId = '', amountMl = '') {
  const container = document.getElementById('preparation-ingredients');
  container.appendChild(createPreparationIngredientRow(ingredientId, amountMl));
}

function renderPreparationsList() {
  const list = document.getElementById('preparations-list');
  const count = document.getElementById('preparations-count');
  if (!list || !count) return;

  count.textContent = state.preparations.length;

  if (state.preparations.length === 0) {
    list.innerHTML = '<p class="list__empty">Пока нет заготовок. Создайте первую!</p>';
    return;
  }

  list.innerHTML = state.preparations
    .map((preparation) => {
      const totalCost = calculatePreparationTotalCost(preparation.items);
      const costPerMl = calculatePreparationCostPerMl(preparation);
      const breakdown = preparation.items
        .map((item) => {
          const ingredient = getIngredientById(item.ingredientId);
          const name = ingredient?.name ?? 'Удалённый ингредиент';
          const itemCost = ingredient
            ? formatCostPerMl(ingredient.bottlePrice, ingredient.bottleVolume) * item.amountMl
            : 0;
          return `
            <li>
              <span>${escapeHtml(name)} <span class="amount">(${item.amountMl} мл)</span></span>
              <span>${formatMoney(itemCost)}</span>
            </li>
          `;
        })
        .join('');

      return `
        <div class="preparation-card" data-id="${preparation.id}">
          <div class="preparation-card__header">
            <div>
              <h3 class="preparation-card__name">${escapeHtml(preparation.name)}</h3>
              <span class="preparation-badge">Заготовка</span>
            </div>
            <span class="preparation-card__per-ml">${formatMoney(costPerMl)}/мл</span>
          </div>
          <p class="preparation-card__meta">
            ${preparation.totalVolume} мл · ${formatMoney(totalCost)} · ${preparation.items.length} ${pluralize(preparation.items.length, ['ингредиент', 'ингредиента', 'ингредиентов'])}
          </p>
          <ul class="cocktail-card__breakdown">${breakdown}</ul>
          <div class="cocktail-card__actions">
            <button class="btn btn--secondary btn--sm btn-edit-preparation" data-id="${preparation.id}">
              ${ICON_EDIT} Редактировать
            </button>
            <button class="btn btn--danger btn--sm btn-delete-preparation" data-id="${preparation.id}">
              ${ICON_DELETE} Удалить
            </button>
          </div>
        </div>
      `;
    })
    .join('');

  list.querySelectorAll('.btn-edit-preparation').forEach((btn) => {
    btn.addEventListener('click', () => editPreparation(btn.dataset.id));
  });

  list.querySelectorAll('.btn-delete-preparation').forEach((btn) => {
    btn.addEventListener('click', () => deletePreparation(btn.dataset.id));
  });
}

function editPreparation(id) {
  const preparation = getPreparationById(id);
  if (!preparation) return;

  state.editingPreparationId = id;
  document.getElementById('preparation-id').value = id;
  document.getElementById('preparation-name').value = preparation.name;
  document.getElementById('preparation-volume').value = preparation.totalVolume;
  document.getElementById('preparation-form-title').textContent = 'Редактировать заготовку';
  document.getElementById('preparation-submit').textContent = 'Сохранить';
  document.getElementById('preparation-cancel').hidden = false;

  const container = document.getElementById('preparation-ingredients');
  container.innerHTML = '';
  preparation.items.forEach((item) => {
    addPreparationIngredientRow(item.ingredientId, item.amountMl);
  });
  if (preparation.items.length === 0) addPreparationIngredientRow();

  updatePreparationPreview();
  document.querySelector('[data-tab="preparations"]').click();
  document.getElementById('preparation-name').focus();
}

function deletePreparation(id) {
  const preparation = getPreparationById(id);
  if (!preparation) return;

  const usedInCocktails = state.cocktails.filter((c) =>
    c.items.some((item) => item.preparationId === id)
  );

  let message = `Удалить заготовку «${preparation.name}»?`;
  if (usedInCocktails.length > 0) {
    message += `\n\nЗаготовка используется в ${usedInCocktails.length} ${pluralize(usedInCocktails.length, ['коктейле', 'коктейлях', 'коктейлях'])}. Она будет удалена из рецептов.`;
  }

  if (!confirm(message)) return;

  state.preparations = state.preparations.filter((p) => p.id !== id);
  state.cocktails = state.cocktails.map((c) => ({
    ...c,
    items: c.items.filter((item) => item.preparationId !== id),
  }));

  if (state.editingPreparationId === id) resetPreparationForm();
  saveData();
  refreshDependentViews();
  showToast('Заготовка удалена');
}

function handlePreparationSubmit(e) {
  e.preventDefault();

  const name = document.getElementById('preparation-name').value.trim();
  const totalVolume = parseFloat(document.getElementById('preparation-volume').value);
  const items = getPreparationFormItems();

  if (!name) {
    showToast('Введите название заготовки', 'error');
    return;
  }
  if (items.length === 0) {
    showToast('Добавьте хотя бы один ингредиент', 'error');
    return;
  }
  if (!totalVolume || totalVolume <= 0) {
    showToast('Укажите итоговый объём заготовки', 'error');
    return;
  }
  if (state.ingredients.length === 0) {
    showToast('Сначала добавьте ингредиенты', 'error');
    return;
  }

  const duplicate = state.preparations.find(
    (p) => p.name.toLowerCase() === name.toLowerCase() && p.id !== state.editingPreparationId
  );
  if (duplicate) {
    showToast('Заготовка с таким названием уже существует', 'error');
    return;
  }

  if (state.editingPreparationId) {
    const index = state.preparations.findIndex((p) => p.id === state.editingPreparationId);
    if (index !== -1) {
      state.preparations[index] = {
        ...state.preparations[index],
        name,
        items,
        totalVolume,
      };
    }
    showToast('Заготовка обновлена');
  } else {
    state.preparations.push({
      id: generateId(),
      name,
      items,
      totalVolume,
    });
    showToast('Заготовка создана');
  }

  saveData();
  resetPreparationForm();
  refreshDependentViews();
}

function initPreparationForm() {
  const form = document.getElementById('preparation-form');
  form.addEventListener('submit', handlePreparationSubmit);
  document.getElementById('add-preparation-ingredient').addEventListener('click', () => {
    addPreparationIngredientRow();
  });
  document.getElementById('preparation-cancel').addEventListener('click', resetPreparationForm);
  document.getElementById('preparation-volume').addEventListener('input', updatePreparationPreview);
}

// ─── Cocktails ───────────────────────────────────────────────────────────────

function buildCocktailItemOptions(selectedRef = '') {
  if (state.ingredients.length === 0 && state.preparations.length === 0) {
    return '<option value="">— Сначала добавьте ингредиенты —</option>';
  }

  let options = '<option value="">— Выберите ингредиент —</option>';

  if (state.ingredients.length > 0) {
    options += '<optgroup label="Ингредиенты">';
    options += state.ingredients
      .map((ing) => {
        const ref = encodeCocktailItemRef('i', ing.id);
        return `<option value="${ref}" ${ref === selectedRef ? 'selected' : ''}>${escapeHtml(ing.name)}</option>`;
      })
      .join('');
    options += '</optgroup>';
  }

  if (state.preparations.length > 0) {
    options += '<optgroup label="Заготовки">';
    options += state.preparations
      .map((prep) => {
        const ref = encodeCocktailItemRef('p', prep.id);
        return `<option value="${ref}" ${ref === selectedRef ? 'selected' : ''}>${escapeHtml(prep.name)}</option>`;
      })
      .join('');
    options += '</optgroup>';
  }

  return options;
}

function createCocktailIngredientRow(itemRef = '', amountMl = '') {
  const selectedRef = itemRef.includes(':')
    ? itemRef
    : itemRef
      ? encodeCocktailItemRef('i', itemRef)
      : '';

  return createCompositionRow({
    selectClass: 'cocktail-ingredient-select',
    amountClass: 'cocktail-ingredient-amount',
    optionsHtml: buildCocktailItemOptions(selectedRef),
    amountMl,
    getCostText: (value, ml) => {
      const refs = parseCocktailItemRef(value);
      const item = { ...refs, amountMl: ml };
      if ((refs.ingredientId || refs.preparationId) && ml > 0) {
        return `= ${formatMoney(getCocktailItemCost(item))}`;
      }
      return '';
    },
    onUpdate: updateCocktailPreview,
  });
}

function refreshCocktailIngredientSelects() {
  const container = document.getElementById('cocktail-ingredients');
  if (!container) return;

  container.querySelectorAll('.cocktail-ingredient-row').forEach((row) => {
    const select = row.querySelector('.cocktail-ingredient-select');
    const current = select.value;
    select.innerHTML = buildCocktailItemOptions(current);
  });
}

function getCocktailFormItems() {
  const rows = document.querySelectorAll('#cocktail-ingredients .cocktail-ingredient-row');
  return Array.from(rows)
    .map((row) => {
      const value = row.querySelector('.cocktail-ingredient-select').value;
      const amountMl = parseFloat(row.querySelector('.cocktail-ingredient-amount').value);
      if (!value || !amountMl || amountMl <= 0) return null;
      return { ...parseCocktailItemRef(value), amountMl };
    })
    .filter(Boolean);
}

function updateCocktailPreview() {
  const items = getCocktailFormItems();
  const preview = document.getElementById('cocktail-preview');
  const valueEl = document.getElementById('cocktail-total-cost');

  if (items.length > 0) {
    preview.hidden = false;
    valueEl.textContent = formatMoney(calculateCocktailCost(items));
  } else {
    preview.hidden = true;
  }
}

function resetCocktailForm() {
  state.editingCocktailId = null;
  document.getElementById('cocktail-id').value = '';
  document.getElementById('cocktail-name').value = '';
  document.getElementById('cocktail-ingredients').innerHTML = '';
  document.getElementById('cocktail-preview').hidden = true;
  document.getElementById('cocktail-form-title').textContent = 'Создать коктейль';
  document.getElementById('cocktail-submit').textContent = 'Создать';
  document.getElementById('cocktail-cancel').hidden = true;
  addCocktailIngredientRow();
}

function addCocktailIngredientRow(itemRef = '', amountMl = '') {
  const container = document.getElementById('cocktail-ingredients');
  container.appendChild(createCocktailIngredientRow(itemRef, amountMl));
}

function renderCocktailsList() {
  const list = document.getElementById('cocktails-list');
  const count = document.getElementById('cocktails-count');
  count.textContent = state.cocktails.length;

  if (state.cocktails.length === 0) {
    list.innerHTML = '<p class="list__empty">Пока нет коктейлей. Создайте первый!</p>';
    return;
  }

  list.innerHTML = state.cocktails
    .map((cocktail) => {
      const total = calculateCocktailCost(cocktail.items);
      const breakdown = cocktail.items
        .map((item) => {
          const name = getCocktailItemName(item);
          const itemCost = getCocktailItemCost(item);
          const badge = isPreparationItem(item)
            ? '<span class="preparation-badge preparation-badge--inline">заготовка</span> '
            : '';
          return `
            <li>
              <span>${escapeHtml(name)} ${badge}<span class="amount">(${item.amountMl} мл)</span></span>
              <span>${formatMoney(itemCost)}</span>
            </li>
          `;
        })
        .join('');

      return `
        <div class="cocktail-card" data-id="${cocktail.id}">
          <div class="cocktail-card__header">
            <h3 class="cocktail-card__name">${escapeHtml(cocktail.name)}</h3>
            <span class="cocktail-card__total">${formatMoney(total)}</span>
          </div>
          <ul class="cocktail-card__breakdown">${breakdown}</ul>
          <div class="cocktail-card__actions">
            <button class="btn btn--secondary btn--sm btn-edit-cocktail" data-id="${cocktail.id}">
              ${ICON_EDIT} Редактировать
            </button>
            <button class="btn btn--danger btn--sm btn-delete-cocktail" data-id="${cocktail.id}">
              ${ICON_DELETE} Удалить
            </button>
          </div>
        </div>
      `;
    })
    .join('');

  list.querySelectorAll('.btn-edit-cocktail').forEach((btn) => {
    btn.addEventListener('click', () => editCocktail(btn.dataset.id));
  });

  list.querySelectorAll('.btn-delete-cocktail').forEach((btn) => {
    btn.addEventListener('click', () => deleteCocktail(btn.dataset.id));
  });
}

function editCocktail(id) {
  const cocktail = state.cocktails.find((c) => c.id === id);
  if (!cocktail) return;

  state.editingCocktailId = id;
  document.getElementById('cocktail-id').value = id;
  document.getElementById('cocktail-name').value = cocktail.name;
  document.getElementById('cocktail-form-title').textContent = 'Редактировать коктейль';
  document.getElementById('cocktail-submit').textContent = 'Сохранить';
  document.getElementById('cocktail-cancel').hidden = false;

  const container = document.getElementById('cocktail-ingredients');
  container.innerHTML = '';
  cocktail.items.forEach((item) => {
    addCocktailIngredientRow(getCocktailItemRef(item), item.amountMl);
  });
  if (cocktail.items.length === 0) addCocktailIngredientRow();

  updateCocktailPreview();

  document.querySelector('[data-tab="cocktails"]').click();
  document.getElementById('cocktail-name').focus();
}

function deleteCocktail(id) {
  const cocktail = state.cocktails.find((c) => c.id === id);
  if (!cocktail) return;
  if (!confirm(`Удалить коктейль «${cocktail.name}»?`)) return;

  state.cocktails = state.cocktails.filter((c) => c.id !== id);
  if (state.editingCocktailId === id) resetCocktailForm();
  saveData();
  renderCocktailsList();
  if (document.getElementById('home-panel')?.classList.contains('panel--active')) {
    renderHomeDashboard();
  }
  showToast('Коктейль удалён');
}

function handleCocktailSubmit(e) {
  e.preventDefault();

  const name = document.getElementById('cocktail-name').value.trim();
  const items = getCocktailFormItems();

  if (!name) {
    showToast('Введите название коктейля', 'error');
    return;
  }
  if (items.length === 0) {
    showToast('Добавьте хотя бы один ингредиент', 'error');
    return;
  }
  if (state.ingredients.length === 0 && state.preparations.length === 0) {
    showToast('Сначала добавьте ингредиенты или заготовки', 'error');
    return;
  }

  const duplicate = state.cocktails.find(
    (c) => c.name.toLowerCase() === name.toLowerCase() && c.id !== state.editingCocktailId
  );
  if (duplicate) {
    showToast('Коктейль с таким названием уже существует', 'error');
    return;
  }

  if (state.editingCocktailId) {
    const index = state.cocktails.findIndex((c) => c.id === state.editingCocktailId);
    if (index !== -1) {
      state.cocktails[index] = { ...state.cocktails[index], name, items };
    }
    showToast('Коктейль обновлён');
  } else {
    state.cocktails.push({ id: generateId(), name, items });
    showToast('Коктейль создан');
  }

  saveData();
  resetCocktailForm();
  renderCocktailsList();
  if (document.getElementById('home-panel')?.classList.contains('panel--active')) {
    renderHomeDashboard();
  }
}

function initCocktailForm() {
  const form = document.getElementById('cocktail-form');
  form.addEventListener('submit', handleCocktailSubmit);
  document.getElementById('add-cocktail-ingredient').addEventListener('click', () => {
    addCocktailIngredientRow();
  });
  document.getElementById('cocktail-cancel').addEventListener('click', resetCocktailForm);
}

// ─── Utils ───────────────────────────────────────────────────────────────────

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

function pluralize(count, [one, few, many]) {
  const abs = Math.abs(count) % 100;
  const rest = abs % 10;
  if (rest === 1 && abs !== 11) return one;
  if (rest >= 2 && rest <= 4 && (abs < 10 || abs >= 20)) return few;
  return many;
}

// ─── PWA ─────────────────────────────────────────────────────────────────────

function registerServiceWorker() {
  if (!('serviceWorker' in navigator)) return;

  window.addEventListener('load', () => {
    navigator.serviceWorker.register('/sw.js').catch(() => {});
  });
}

// ─── Init ────────────────────────────────────────────────────────────────────

function init() {
  loadData();
  initTabs();
  initIngredientForm();
  initPreparationForm();
  initCocktailForm();
  initBackupImport();
  renderHomeDashboard();
  renderIngredientsList();
  renderPreparationsList();
  renderCocktailsList();
  resetPreparationForm();
  resetCocktailForm();
  registerServiceWorker();
}

init();
