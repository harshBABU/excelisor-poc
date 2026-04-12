import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional, Union
import logging
from sklearn.preprocessing import StandardScaler
from scipy import stats
import random

logger = logging.getLogger(__name__)

class IntelligentSampler:
    def __init__(self, max_sample_size: int = 10000, min_sample_size: int = 2):
        self.max_sample_size = max_sample_size
        self.min_sample_size = min_sample_size
        
    def sample_large_dataset(self, data: Union[pd.DataFrame, List[List[Any]]]) -> Dict[str, Any]:
        """
        Multi-stage intelligent sampling for datasets with 1M+ rows
        Accepts either a pandas DataFrame or List[List[Any]]
        """
        try:
            # Handle both DataFrame and list input
            if isinstance(data, pd.DataFrame):
                if data.empty:
                    return {"error": "No data provided"}
                df = data
            elif isinstance(data, list):
                if not data or len(data) == 0:
                    return {"error": "No data provided"}
                df = pd.DataFrame(data)
                
            else:
                return {"error": "No data provided"}
            #print(f"Data: {data}")
            #print(f"DataFrame: {df}")
            total_rows = len(df)
            
            print(f"Processing dataset with {total_rows:,} rows")
            
            # Always analyze headers
            header_analysis = self.analyze_headers(df)
            
            if total_rows <= self.max_sample_size:
                print(f"Small dataset - returning all data")
                # Include headers as first row, then data rows
                pattern_data = [df.columns.tolist()] + df.values.tolist()
                return {
                    "header_analysis": header_analysis,
                    "pattern_sample": pattern_data,
                    "edge_cases": self.detect_edge_cases(df),
                    "metadata": self.generate_metadata(df),
                    "full_data_available": False,
                    "sampling_strategy": "full_data"
                }
            
            # Large dataset - intelligent sampling
            sample_data = self.pattern_based_sample(df)
            
            # Include headers as first row, then sampled data rows
            pattern_data = [df.columns.tolist()] + sample_data.values.tolist()
            return {
                "header_analysis": header_analysis,
                "pattern_sample": pattern_data,
                "edge_cases": self.detect_edge_cases(df, sample_size=5000),
                "metadata": self.generate_metadata(sample_data, full_df=df),
                "full_data_available": True,
                "sampling_strategy": "intelligent_multi_stage",
                "original_size": total_rows,
                "sample_size": len(sample_data)
            }
            
        except Exception as e:
            logger.error(f"Sampling error: {str(e)}")
            return {"error": f"Sampling failed: {str(e)}"}
    
    def analyze_headers(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Analyze column headers and data types"""
        try:
            from datetime import datetime as _dt, timedelta as _td

            # Excel epoch for serial-date detection
            _EXCEL_EPOCH  = _dt(1899, 12, 30)
            _SERIAL_MIN   = 25569   # 1970-01-01
            _SERIAL_MAX   = 73050   # 2099-12-31
            _DATE_KEYWORDS = ("date", "time", "dt", "day", "month", "year",
                              "period", "created", "updated")

            def _looks_like_excel_date_col(col_name: str, series: pd.Series) -> bool:
                """Heuristic: numeric column with a date-like name whose values
                sit inside the Excel serial-number range."""
                if not any(kw in col_name.lower() for kw in _DATE_KEYWORDS):
                    return False
                if not pd.api.types.is_numeric_dtype(series):
                    return False
                vals = series.dropna()
                if len(vals) == 0:
                    return False
                try:
                    return bool(
                        (vals >= _SERIAL_MIN).all() and (vals <= _SERIAL_MAX).all()
                    )
                except Exception:
                    return False

            analysis = {}
            for col in df.columns:
                col_data = df[col].dropna()
                if len(col_data) == 0:
                    continue
                    
                # Detect data type and patterns
                is_excel_dt = _looks_like_excel_date_col(str(col), df[col])

                analysis[str(col)] = {
                    "dtype": str(df[col].dtype),
                    "null_count": df[col].isnull().sum(),
                    "null_percentage": (df[col].isnull().sum() / len(df)) * 100,
                    "unique_values": df[col].nunique(),
                    "sample_values": col_data.head(5).tolist(),
                    "is_numeric": pd.api.types.is_numeric_dtype(df[col]) and not is_excel_dt,
                    "is_datetime": pd.api.types.is_datetime64_any_dtype(df[col]),
                    "is_date": is_excel_dt or pd.api.types.is_datetime64_any_dtype(df[col]),
                    "is_categorical": self.is_likely_categorical(df[col])
                }
                
                # Additional analysis for numeric columns (skip if it's an Excel date serial)
                if analysis[str(col)]["is_numeric"] and not is_excel_dt:
                    analysis[str(col)].update({
                        "min": col_data.min(),
                        "max": col_data.max(),
                        "mean": col_data.mean(),
                        "std": col_data.std()
                    })
            
            return analysis
            
        except Exception as e:
            logger.error(f"Header analysis error: {str(e)}")
            return {}
    
    def pattern_based_sample(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Advanced sampling strategy combining multiple methods:
        - 20% systematic sampling (every nth row)
        - 60% stratified sampling (based on key patterns)
        - 20% random sampling
        """
        total_rows = len(df)
        target_size = min(self.max_sample_size, max(self.min_sample_size, int(total_rows * 0.01)))
        
        samples = []
        
        # 1. Systematic sampling (20%)
        systematic_size = int(target_size * 0.2)
        step = total_rows // systematic_size
        systematic_indices = list(range(0, total_rows, step))[:systematic_size]
        samples.extend(systematic_indices)
        
        # 2. Stratified sampling (60%) - based on data patterns
        stratified_size = int(target_size * 0.6)
        stratified_indices = self.stratified_sample_indices(df, stratified_size)
        samples.extend(stratified_indices)
        
        # 3. Random sampling (20%)
        random_size = target_size - len(samples)
        remaining_indices = list(set(range(total_rows)) - set(samples))
        if remaining_indices and random_size > 0:
            random_indices = random.sample(remaining_indices, min(random_size, len(remaining_indices)))
            samples.extend(random_indices)
        
        # Remove duplicates and sort
        final_indices = sorted(list(set(samples)))
        
        return df.iloc[final_indices]
    
    def stratified_sample_indices(self, df: pd.DataFrame, sample_size: int) -> List[int]:
        """Create stratified sample based on data distribution"""
        try:
            # Find the best column for stratification
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            categorical_cols = df.select_dtypes(include=['object', 'category']).columns
            
            stratify_col = None
            
            # Prefer categorical columns with reasonable number of categories
            for col in categorical_cols:
                unique_count = df[col].nunique()
                if 2 <= unique_count <= 50:  # Reasonable number of categories
                    stratify_col = col
                    break
            
            # Fallback to numeric column quartiles
            if stratify_col is None and len(numeric_cols) > 0:
                col = numeric_cols[0]
                df_temp = df.copy()
                df_temp['quartile'] = pd.qcut(df_temp[col], q=4, duplicates='drop')
                stratify_col = 'quartile'
            
            if stratify_col is None:
                # Fallback to random sampling
                return random.sample(range(len(df)), min(sample_size, len(df)))
            
            # Perform stratified sampling
            groups = df.groupby(stratify_col)
            samples_per_group = max(1, sample_size // len(groups))
            
            indices = []
            for name, group in groups:
                group_sample_size = min(samples_per_group, len(group))
                group_indices = group.index.tolist()
                sampled_indices = random.sample(group_indices, group_sample_size)
                indices.extend(sampled_indices)
            
            return indices
            
        except Exception as e:
            logger.error(f"Stratified sampling error: {str(e)}")
            return random.sample(range(len(df)), min(sample_size, len(df)))
    
    def detect_edge_cases(self, df: pd.DataFrame, sample_size: int = 1000) -> Dict[str, Any]:
        """Detect outliers and edge cases in the dataset"""
        try:
            # Sample for edge case detection
            if len(df) > sample_size:
                sample_df = df.sample(n=sample_size)
            else:
                sample_df = df
            
            edge_cases = {}
            numeric_cols = sample_df.select_dtypes(include=[np.number]).columns
            
            for col in numeric_cols:
                col_data = sample_df[col].dropna()
                if len(col_data) == 0:
                    continue
                
                # Use IQR method for outlier detection
                Q1 = col_data.quantile(0.25)
                Q3 = col_data.quantile(0.75)
                IQR = Q3 - Q1
                lower_bound = Q1 - 1.5 * IQR
                upper_bound = Q3 + 1.5 * IQR
                
                outliers = col_data[(col_data < lower_bound) | (col_data > upper_bound)]
                
                edge_cases[str(col)] = {
                    "outlier_count": len(outliers),
                    "outlier_percentage": (len(outliers) / len(col_data)) * 100,
                    "outlier_values": outliers.head(10).tolist(),
                    "bounds": {"lower": lower_bound, "upper": upper_bound}
                }
            
            return edge_cases
            
        except Exception as e:
            logger.error(f"Edge case detection error: {str(e)}")
            return {}
    
    def generate_metadata(self, sample_df: pd.DataFrame, full_df: Optional[pd.DataFrame] = None) -> Dict[str, Any]:
        """Generate comprehensive dataset metadata"""
        try:
            df_to_analyze = full_df if full_df is not None else sample_df
            
            metadata = {
                "shape": {
                    "total_rows": len(df_to_analyze),
                    "total_columns": len(df_to_analyze.columns),
                    "sample_rows": len(sample_df)
                },
                "data_types": {
                    "numeric_columns": len(df_to_analyze.select_dtypes(include=[np.number]).columns),
                    "text_columns": len(df_to_analyze.select_dtypes(include=['object']).columns),
                    "datetime_columns": len(df_to_analyze.select_dtypes(include=['datetime']).columns)
                },
                "data_quality": {
                    "total_nulls": df_to_analyze.isnull().sum().sum(),
                    "null_percentage": (df_to_analyze.isnull().sum().sum() / df_to_analyze.size) * 100,
                    "duplicate_rows": df_to_analyze.duplicated().sum()
                },
                "patterns": self.detect_data_patterns(sample_df),
                "complexity_score": self.calculate_complexity_score(sample_df)
            }
            
            return metadata
            
        except Exception as e:
            logger.error(f"Metadata generation error: {str(e)}")
            return {}
    
    def detect_data_patterns(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Detect common data patterns"""
        patterns = {
            "has_time_series": False,
            "has_hierarchical": False,
            "has_categories": False,
            "has_metrics": False
        }
        
        try:
            # Check for time series patterns
            datetime_cols = df.select_dtypes(include=['datetime']).columns
            if len(datetime_cols) > 0:
                patterns["has_time_series"] = True
            
            # Check for categorical data
            categorical_cols = df.select_dtypes(include=['object', 'category']).columns
            if len(categorical_cols) > 0:
                patterns["has_categories"] = True
                
            # Check for numeric metrics
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            if len(numeric_cols) > 0:
                patterns["has_metrics"] = True
            
            # Check for hierarchical patterns (columns with common prefixes)
            col_names = [str(col) for col in df.columns]
            prefixes = {}
            for col in col_names:
                parts = col.split('_')
                if len(parts) > 1:
                    prefix = parts[0]
                    prefixes[prefix] = prefixes.get(prefix, 0) + 1
            
            if any(count > 2 for count in prefixes.values()):
                patterns["has_hierarchical"] = True
                
        except Exception as e:
            logger.error(f"Pattern detection error: {str(e)}")
            
        return patterns
    
    def calculate_complexity_score(self, df: pd.DataFrame) -> str:
        """Calculate dataset complexity for AI model selection"""
        try:
            score = 0
            
            # Size factor
            if len(df) > 100000:
                score += 3
            elif len(df) > 10000:
                score += 2
            else:
                score += 1
            
            # Column diversity
            if len(df.columns) > 20:
                score += 3
            elif len(df.columns) > 10:
                score += 2
            else:
                score += 1
            
            # Data type diversity
            dtypes = len(df.dtypes.unique())
            if dtypes > 3:
                score += 2
            elif dtypes > 1:
                score += 1
            
            # Missing data complexity
            null_percentage = (df.isnull().sum().sum() / df.size) * 100
            if null_percentage > 20:
                score += 2
            elif null_percentage > 5:
                score += 1
            
            # Return complexity level
            if score >= 8:
                return "expert"
            elif score >= 5:
                return "complex"
            elif score >= 3:
                return "medium"
            else:
                return "simple"
                
        except Exception:
            return "medium"
    
    def is_likely_categorical(self, series: pd.Series) -> bool:
        """Determine if a column is likely categorical"""
        # Check column name for obvious categorical indicators (IDs, codes, etc.)
        col_name = str(series.name).lower() if series.name else ""
        categorical_name_patterns = ['id', 'code', 'sku', 'product', 'customer', 'user', 'category', 'type', 'class', 'group', 'channel', 'warehouse']
        
        # If column name suggests it's categorical, treat it as such regardless of dtype
        if any(pattern in col_name for pattern in categorical_name_patterns):
            unique_ratio = series.nunique() / len(series)
            # More lenient for ID/code columns - they can have higher unique ratios
            return unique_ratio < 0.8 or series.nunique() < 1000
        
        # Original logic for other columns
        if series.dtype == 'object' or series.dtype.name == 'category':
            unique_ratio = series.nunique() / len(series)
            return unique_ratio < 0.1 or series.nunique() < 50
        
        return False
    
    def calculate_quality_score(self, df: pd.DataFrame) -> float:
        """Calculate overall data quality score (0-10)"""
        try:
            score = 10.0
            
            # Penalize missing values
            null_percentage = (df.isnull().sum().sum() / df.size) * 100
            score -= min(null_percentage / 10, 3)  # Max 3 points deduction
            
            # Penalize duplicate rows
            dup_percentage = (df.duplicated().sum() / len(df)) * 100
            score -= min(dup_percentage / 10, 2)  # Max 2 points deduction
            
            # Check for data consistency
            for col in df.select_dtypes(include=[np.number]).columns:
                col_data = df[col].dropna()
                if len(col_data) > 0:
                    # Check for extreme outliers
                    z_scores = np.abs(stats.zscore(col_data))
                    extreme_outliers = (z_scores > 4).sum()
                    outlier_percentage = (extreme_outliers / len(col_data)) * 100
                    score -= min(outlier_percentage / 5, 1)  # Max 1 point per column
            
            return max(0, min(10, score))
            
        except Exception:
            return 7.0  # Default score
