import os
import json
try:
    import mysql.connector
except ImportError:
    mysql = None

from sqlite3 import connect as sqlite_connect
from dotenv import load_dotenv

load_dotenv()

class DatabaseManager:
    def __init__(self):
        self.db_type = os.getenv("DB_TYPE", "sqlite")
        self.host = os.getenv("DB_HOST", "localhost")
        self.user = os.getenv("DB_USER", "root")
        self.password = os.getenv("DB_PASSWORD", "")
        self.database = os.getenv("DB_NAME", "retail_roi")
        self.conn = None
        
    def get_connection(self):
        if self.db_type == "mysql" and mysql:
            try:
                return mysql.connector.connect(
                    host=self.host,
                    user=self.user,
                    password=self.password,
                    database=self.database
                )
            except Exception as e:
                print(f"MySQL Error: {e}. Falling back to SQLite.")
                self.db_type = "sqlite"
                return sqlite_connect("retail_roi.db")
        else:
            return sqlite_connect("retail_roi.db")

    def init_db(self):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        is_mysql = self.db_type == "mysql"
        pk_auto = "AUTO_INCREMENT" if is_mysql else "AUTOINCREMENT"
        
        # 1. Modules/Config Tables (Master Configuration)
        cursor.execute(f"CREATE TABLE IF NOT EXISTS modules (id INTEGER PRIMARY KEY {pk_auto if not is_mysql else 'AUTO_INCREMENT'}, name VARCHAR(255) UNIQUE)")
        cursor.execute(f"CREATE TABLE IF NOT EXISTS module_aspects (id INTEGER PRIMARY KEY {pk_auto if not is_mysql else 'AUTO_INCREMENT'}, module_name VARCHAR(255), aspect_key VARCHAR(100))")
        cursor.execute(f"CREATE TABLE IF NOT EXISTS benefit_params (id INTEGER PRIMARY KEY {pk_auto if not is_mysql else 'AUTO_INCREMENT'}, module_name VARCHAR(255), aspect_key VARCHAR(100), min_val FLOAT, max_val FLOAT)")
        cursor.execute(f"CREATE TABLE IF NOT EXISTS annual_investments (year_val INTEGER PRIMARY KEY, software FLOAT, impl FLOAT, extra FLOAT)")
        cursor.execute(f"CREATE TABLE IF NOT EXISTS aspect_ranges (aspect_key VARCHAR(100) PRIMARY KEY, min_val FLOAT, max_val FLOAT)")
        
        # 2. Exercise Persistence Tables (ROI Runs)
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS exercises (
                id INTEGER PRIMARY KEY {pk_auto if not is_mysql else 'AUTO_INCREMENT'},
                exercise_name VARCHAR(255),
                client_name VARCHAR(255),
                retailer_type VARCHAR(100),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                net_revenue FLOAT,
                growth_rate FLOAT,
                inventory FLOAT,
                carrying_cost FLOAT,
                cogs_pct FLOAT,
                sga_pct FLOAT,
                tax_rate FLOAT,
                discount_rate FLOAT,
                adoption_years INTEGER,
                scenario_type VARCHAR(50)
            )
        """)
        
        cursor.execute(f"CREATE TABLE IF NOT EXISTS exercise_modules (exercise_id INTEGER, module_name VARCHAR(255))")
        cursor.execute(f"CREATE TABLE IF NOT EXISTS exercise_benefits (exercise_id INTEGER, module_name VARCHAR(255), aspect_key VARCHAR(100), percentage FLOAT)")
        cursor.execute(f"CREATE TABLE IF NOT EXISTS exercise_investments (exercise_id INTEGER, year_val INTEGER, software FLOAT, impl FLOAT, extra FLOAT)")

        conn.commit()
        conn.close()

    # --- Master Config Methods ---
    def load_modules(self):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM modules")
        rows = cursor.fetchall()
        modules = [r[0] for r in rows]
        conn.close()
        return modules

    def load_profiles(self):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT module_name, aspect_key FROM module_aspects")
        rows = cursor.fetchall()
        profiles = {}
        for mod, aspect in rows:
            if mod not in profiles: profiles[mod] = []
            profiles[mod].append(aspect)
        conn.close()
        return profiles

    def load_benefit_params(self):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT module_name, aspect_key, min_val, max_val FROM benefit_params")
        rows = cursor.fetchall()
        params = {}
        for mod, aspect, min_v, max_v in rows:
            if mod not in params: params[mod] = {}
            params[mod][aspect] = {"min": min_v, "max": max_v}
        conn.close()
        return params

    def load_investments(self):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT year_val, software, impl, extra FROM annual_investments")
        rows = cursor.fetchall()
        investments = {r[0]: {"software": r[1], "impl": r[2], "extra": r[3]} for r in rows}
        conn.close()
        return investments

    def load_aspect_ranges(self):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT aspect_key, min_val, max_val FROM aspect_ranges")
        rows = cursor.fetchall()
        ranges = {r[0]: (r[1], r[2]) for r in rows}
        conn.close()
        return ranges

    def sync_all(self, modules, profiles, benefit_params, investments, aspect_ranges):
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("DELETE FROM modules")
            for m in modules:
                cursor.execute("INSERT INTO modules (name) VALUES (%s)" if self.db_type=="mysql" else "INSERT INTO modules (name) VALUES (?)", (m,))
            
            cursor.execute("DELETE FROM module_aspects")
            for mod, aspects in profiles.items():
                for a in aspects:
                    cursor.execute("INSERT INTO module_aspects (module_name, aspect_key) VALUES (%s, %s)" if self.db_type=="mysql" else "INSERT INTO module_aspects (module_name, aspect_key) VALUES (?, ?)", (mod, a))
            
            cursor.execute("DELETE FROM benefit_params")
            for mod, aspects in benefit_params.items():
                for aspect, vals in aspects.items():
                    cursor.execute("INSERT INTO benefit_params (module_name, aspect_key, min_val, max_val) VALUES (%s, %s, %s, %s)" if self.db_type=="mysql" else "INSERT INTO benefit_params (module_name, aspect_key, min_val, max_val) VALUES (?, ?, ?, ?)", (mod, aspect, vals['min'], vals['max']))
            
            cursor.execute("DELETE FROM annual_investments")
            for year, vals in investments.items():
                cursor.execute("INSERT INTO annual_investments (year_val, software, impl, extra) VALUES (%s, %s, %s, %s)" if self.db_type=="mysql" else "INSERT INTO annual_investments (year_val, software, impl, extra) VALUES (?, ?, ?, ?)", (year, vals['software'], vals['impl'], vals['extra']))

            cursor.execute("DELETE FROM aspect_ranges")
            for key, (min_v, max_v) in aspect_ranges.items():
                cursor.execute("INSERT INTO aspect_ranges (aspect_key, min_val, max_val) VALUES (%s, %s, %s)" if self.db_type=="mysql" else "INSERT INTO aspect_ranges (aspect_key, min_val, max_val) VALUES (?, ?, ?)", (key, min_v, max_v))
            
            conn.commit()
        except Exception as e:
            print(f"Sync error: {e}")
        finally:
            conn.close()

    # --- Exercise Methods ---
    def save_exercise(self, data):
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            placeholder = "%s" if self.db_type == "mysql" else "?"
            
            # Insert main exercise record
            cursor.execute(f"""
                INSERT INTO exercises (exercise_name, client_name, retailer_type, net_revenue, growth_rate, 
                inventory, carrying_cost, cogs_pct, sga_pct, tax_rate, discount_rate, adoption_years, scenario_type)
                VALUES ({placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder}, 
                        {placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder}, 
                        {placeholder}, {placeholder}, {placeholder})
            """, (
                data['exercise_name'], data['client_name'], data['retailer_type'], data['net_revenue'],
                data['growth_rate'], data['inventory'], data['carrying_cost'], data['cogs_pct'],
                data['sga_pct'], data['tax_rate'], data['discount_rate'], data['adoption_years'], data['scenario_type']
            ))
            
            exercise_id = cursor.lastrowid
            
            # Insert modules
            for mod in data['module_selected']:
                cursor.execute(f"INSERT INTO exercise_modules (exercise_id, module_name) VALUES ({placeholder}, {placeholder})", (exercise_id, mod))
            
            # Insert benefits
            for mod_name, benefits in data['module_benefits'].items():
                for aspect, pct in benefits.items():
                    cursor.execute(f"INSERT INTO exercise_benefits (exercise_id, module_name, aspect_key, percentage) VALUES ({placeholder}, {placeholder}, {placeholder}, {placeholder})", (exercise_id, mod_name, aspect, pct))
            
            # Insert investments
            for year, vals in data['annual_investments'].items():
                cursor.execute(f"INSERT INTO exercise_investments (exercise_id, year_val, software, impl, extra) VALUES ({placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder})", (exercise_id, year, vals['software'], vals['impl'], vals['extra']))
            
            conn.commit()
            return exercise_id
        except Exception as e:
            print(f"Save exercise error: {e}")
            return None
        finally:
            conn.close()

    def get_exercise_list(self):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id, exercise_name, client_name, created_at FROM exercises ORDER BY created_at DESC")
        rows = cursor.fetchall()
        conn.close()
        return rows

    def load_exercise(self, e_id):
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            placeholder = "%s" if self.db_type == "mysql" else "?"
            
            # Load master
            cursor.execute(f"SELECT * FROM exercises WHERE id = {placeholder}", (e_id,))
            cols = [d[0] for d in cursor.description]
            row = cursor.fetchone()
            if not row: return None
            exercise = dict(zip(cols, row))
            
            # Load modules
            cursor.execute(f"SELECT module_name FROM exercise_modules WHERE exercise_id = {placeholder}", (e_id,))
            exercise['module_selected'] = [r[0] for r in cursor.fetchall()]
            
            # Load benefits
            cursor.execute(f"SELECT module_name, aspect_key, percentage FROM exercise_benefits WHERE exercise_id = {placeholder}", (e_id,))
            benefits = {}
            for mod, aspect, pct in cursor.fetchall():
                if mod not in benefits: benefits[mod] = {}
                benefits[mod][aspect] = pct
            exercise['module_benefits'] = benefits
            
            # Load investments
            cursor.execute(f"SELECT year_val, software, impl, extra FROM exercise_investments WHERE exercise_id = {placeholder}", (e_id,))
            investments = {r[0]: {"software": r[1], "impl": r[2], "extra": r[3]} for r in cursor.fetchall()}
            exercise['annual_investments'] = investments
            
            return exercise
        finally:
            conn.close()

    def delete_exercise(self, e_id):
        conn = self.get_connection()
        cursor = conn.cursor()
        placeholder = "%s" if self.db_type == "mysql" else "?"
        cursor.execute(f"DELETE FROM exercises WHERE id = {placeholder}", (e_id,))
        cursor.execute(f"DELETE FROM exercise_modules WHERE exercise_id = {placeholder}", (e_id,))
        cursor.execute(f"DELETE FROM exercise_benefits WHERE exercise_id = {placeholder}", (e_id,))
        cursor.execute(f"DELETE FROM exercise_investments WHERE exercise_id = {placeholder}", (e_id,))
        conn.commit()
        conn.close()
