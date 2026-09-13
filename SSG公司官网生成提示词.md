# Role
你是一位精通 Nuxt 3、Vue 3 和现代前端架构的全栈开发专家。你对 SSG（静态站点生成）、SEO 优化和服务器端渲染（SSR）水合机制有深入理解。

# Goal
请帮我初始化一个名为 "GS Tech" 的企业官网项目。该项目必须使用 Nuxt 3 的 SSG 模式（静态站点生成）。

# Tech Stack & Core Requirements
1. **框架**: Nuxt 3 (最新版)，使用 TypeScript。
2. **样式**: Tailwind CSS (用于快速构建响应式布局)。
3. **渲染模式**: SSG (Static Site Generation)。
   - 所有页面必须在构建时预渲染为静态 HTML。
   - 必须处理动态路由（如产品详情、新闻详情）的预渲染配置。
4. **数据模拟**:
   - 请在 `server/data/` 目录下创建本地 JSON 文件来模拟数据库内容（产品列表、新闻列表）。
   - 使用 Nuxt 的 Server API Routes (`server/api/`) 来提供数据接口，模拟真实开发环境。
5. **交互需求**:
   - 页面主体内容静态生成。
   - "点击交互加载数据"（如筛选、加载更多、评论提交）必须作为客户端行为，使用 `<ClientOnly>` 包裹或通过 `useFetch` 在客户端调用。

# Pages Structure
请生成以下页面的完整代码结构：

1. **首页 (`pages/index.vue`)**:
   - Hero Banner 区域。
   - "精选产品"板块（从 API 获取前 3 个产品）。
   - "最新新闻"板块（从 API 获取最新 3 条新闻）。
   - 必须配置 SEO Meta 标签。

2. **产品中心 (`pages/products/index.vue`)**:
   - 展示所有产品列表。
   - 支持客户端分页或"加载更多"按钮（演示客户端交互）。

3. **产品详情 (`pages/products/[id].vue`)**:
   - 动态路由页面。
   - 根据 ID 获取产品详细信息。
   - 必须在 `nuxt.config.ts` 中配置此动态路由的预渲染规则。

4. **新闻中心 (`pages/news/index.vue`)**:
   - 新闻列表展示。

5. **新闻详情 (`pages/news/[id].vue`)**:
   - 动态路由页面，展示文章详情。

# Critical SSG Implementation (关键考点)
请在 `nuxt.config.ts` 中编写具体的 SSG 配置逻辑：
- 使用 `nitro.prerender` 配置。
- 编写 Hook (`hooks: { 'prerender:routes' }`)，模拟在构建时从 API 获取所有产品和新闻的 ID，并将它们添加到预渲染路由列表中，确保 `/products/1`, `/products/2` 等页面能生成静态 HTML。

# Output Steps
请按照以下顺序输出代码和解释：

1. **初始化命令**: 给出创建项目和安装 Tailwind CSS 的命令。
2. **目录结构**: 列出关键的文件结构树。
3. **核心配置**: 完整的 `nuxt.config.ts` 代码（重点展示 SSG 和预渲染配置）。
4. **模拟数据**: `server/data/` 和 `server/api/` 的实现代码。
5. **布局与页面**:
   - `layouts/default.vue` (包含导航栏和页脚)。
   - `pages/index.vue`。
   - `pages/products/index.vue` 和 `pages/products/[id].vue`。
   - `pages/news/` 相关页面。
6. **组件逻辑**: 演示如何在页面中混合使用 SSG 数据（首屏）和 CSR 数据（点击交互）。

请确认理解后开始生成代码。
