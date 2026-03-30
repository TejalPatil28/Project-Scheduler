const API = {
  async req(method, path, body) {
    const opts = {
      method,
      credentials: "include",
      headers: { "Content-Type": "application/json" },
    };
    if (body !== undefined) opts.body = JSON.stringify(body);
    const res = await fetch("/api" + path, opts);
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw { status: res.status, message: data.error || "Request failed" };
    return data;
  },

  login:      (username, password) => API.req("POST", "/auth/login",  { username, password }),
  logout:     ()                   => API.req("POST", "/auth/logout"),
  me:         ()                   => API.req("GET",  "/auth/me"),

  getProjects:    ()           => API.req("GET",  "/projects"),
  getProject:     (id)         => API.req("GET",  `/projects/${id}`),
  getTasks:       (id)         => API.req("GET",  `/projects/${id}/tasks`),
  getSheet:       (id)         => API.req("GET",  `/projects/${id}/sheet`),
  saveTasks:      (id, updates)=> API.req("PUT",  `/projects/${id}/tasks`, updates),
  setAssignments: (id, data)   => API.req("PUT",  `/projects/${id}/assignments`, data),
  updateMonitorTimestamp: (id, timestamp) => API.req("POST", `/projects/${id}/monitor-timestamp`, { timestamp }),

  getUsers:   ()           => API.req("GET",    "/users"),
  createUser: (data)       => API.req("POST",   "/users", data),
  updateUser: (username, data) => API.req("PUT", `/users/${username}`, data),
  deleteUser: (username)   => API.req("DELETE", `/users/${username}`),
};
