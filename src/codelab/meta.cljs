(ns codelab.meta
  "FrontMatter (YAML header) parser for Codelab markdown files"
  (:require [clojure.string :as str]))

(def ^:private frontmatter-pattern
  "Regex to extract frontmatter block between --- markers"
  #"(?s)^---\r?\n(.*?)\r?\n---\r?\n?")

(defn- parse-value
  "Parse a YAML value, handling basic types"
  [v]
  (cond
    (nil? v) nil
    (re-matches #"^\d+$" v) (js/parseInt v 10)
    (re-matches #"^(true|false)$" v) (= v "true")
    :else (str/trim v)))

(defn- normalize-key
  "Normalize key to snake_case for consistency with claat"
  [k]
  (-> k
      str/trim
      str/lower-case
      (str/replace #"\s+" "_")))

(defn- parse-line
  "Parse a single key: value line"
  [line]
  (when-let [[_ k v] (re-matches #"^([^:]+):\s*(.*)$" line)]
    [(normalize-key k) (parse-value v)]))

(defn parse-frontmatter
  "Extract and parse frontmatter from markdown content.
   Returns {:meta {...} :content \"...\"} where content is the markdown without frontmatter."
  [content]
  (if-let [[match yaml-content] (re-find frontmatter-pattern content)]
    (let [lines (str/split-lines yaml-content)
          meta (->> lines
                    (map parse-line)
                    (filter some?)
                    (into {}))]
      {:meta meta
       :content (subs content (count match))})
    {:meta {}
     :content content}))

(defn get-meta
  "Get a metadata value with fallback"
  [meta key & [default]]
  (get meta key default))

;; Codelab metadata keys
(def codelab-keys
  "Reserved metadata keys for Codelab"
  #{"id" "url" "summary" "author" "authors"
    "category" "categories"
    "environment" "environments" "tags"
    "status" "state"
    "feedback" "feedback_link"
    "analytics" "analytics_account" "google_analytics"})
