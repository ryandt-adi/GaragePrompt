"""
services.py
Business logic services for Analog Garage Workbench.
Handles prompt generation and validation logic.
"""

from typing import Dict, Optional, Tuple
from dataclasses import dataclass

from templates import template_registry, PromptTemplate
from config import FORM_DEFAULTS


@dataclass
class ValidationResult:
    """Result of form validation."""
    is_valid: bool
    errors: Dict[str, str]
    warnings: Dict[str, str]
    
    @property
    def has_warnings(self) -> bool:
        return len(self.warnings) > 0


class PromptGeneratorService:
    """
    Service for generating prompts from templates.
    Handles validation, context building, and prompt rendering.
    """
    
    def __init__(self, default_template_id: str = "value_creation_v1"):
        self.default_template_id = default_template_id
    
    def validate_context(
        self, 
        context: Dict[str, str], 
        template_id: Optional[str] = None
    ) -> ValidationResult:
        """
        Validate the context against template requirements.
        
        Args:
            context: Dictionary of form field values
            template_id: Optional template ID (uses default if not provided)
        
        Returns:
            ValidationResult with errors and warnings
        """
        errors = {}
        warnings = {}
        
        template = template_registry.get(template_id or self.default_template_id)
        if not template:
            errors["template"] = f"Template not found: {template_id}"
            return ValidationResult(is_valid=False, errors=errors, warnings=warnings)
        
        # Check required fields
        for field in template.required_fields:
            value = context.get(field, "").strip()
            if not value:
                errors[field] = f"{self._format_field_name(field)} is required"
        
        # Check optional fields for sensible values
        for field in template.optional_fields:
            value = context.get(field, "")
            if not value:
                warnings[field] = f"{self._format_field_name(field)} not provided, using default"
        
        # Business-specific validations
        if "innovation_description" in context:
            desc = context["innovation_description"].strip()
            if len(desc) < 20:
                warnings["innovation_description"] = "Description is very short; consider adding more detail"
            elif len(desc) > 2000:
                warnings["innovation_description"] = "Description is very long; consider being more concise"
        
        is_valid = len(errors) == 0
        return ValidationResult(is_valid=is_valid, errors=errors, warnings=warnings)
    
    def build_context(
        self,
        innovation_name: str,
        innovation_description: str,
        industry: str,
        geographic_scope: Optional[str] = None,
        analysis_timeframe: Optional[str] = None,
        innovation_stage: Optional[str] = None,
        currency: Optional[str] = None,
        **extra_fields
    ) -> Dict[str, str]:
        """
        Build a complete context dictionary with defaults.
        
        Args:
            innovation_name: Name of the innovation
            innovation_description: Description of the innovation
            industry: Target industry
            geographic_scope: Geographic scope (optional)
            analysis_timeframe: Analysis timeframe (optional)
            innovation_stage: Stage of innovation (optional)
            currency: Currency for monetary values (optional)
            **extra_fields: Any additional fields
        
        Returns:
            Complete context dictionary
        """
        context = {
            "innovation_name": innovation_name.strip(),
            "innovation_description": innovation_description.strip(),
            "industry": industry.strip(),
            "geographic_scope": geographic_scope or FORM_DEFAULTS["geographic_scope"],
            "analysis_timeframe": analysis_timeframe or FORM_DEFAULTS["analysis_timeframe"],
            "innovation_stage": innovation_stage or FORM_DEFAULTS["innovation_stage"],
            "currency": currency or FORM_DEFAULTS["currency"],
        }
        
        # Add any extra fields
        context.update(extra_fields)
        
        return context
    
    def generate_prompt(
        self, 
        context: Dict[str, str], 
        template_id: Optional[str] = None,
        validate: bool = True
    ) -> Tuple[str, ValidationResult]:
        """
        Generate a prompt from context.
        
        Args:
            context: Dictionary of form field values
            template_id: Optional template ID
            validate: Whether to validate before generating
        
        Returns:
            Tuple of (generated_prompt, validation_result)
        """
        template_id = template_id or self.default_template_id
        
        # Validate if requested
        if validate:
            validation = self.validate_context(context, template_id)
            if not validation.is_valid:
                return "", validation
        else:
            validation = ValidationResult(is_valid=True, errors={}, warnings={})
        
        # Get template and render
        template = template_registry.get(template_id)
        if not template:
            validation.errors["template"] = f"Template not found: {template_id}"
            validation.is_valid = False
            return "", validation
        
        try:
            prompt = template.render(context)
            return prompt, validation
        except ValueError as e:
            validation.errors["render"] = str(e)
            validation.is_valid = False
            return "", validation
    
    def get_available_templates(self) -> list:
        """Get list of available templates."""
        return [
            {
                "id": t.id,
                "name": t.name,
                "description": t.description,
                "category": t.category,
            }
            for t in template_registry.get_all()
        ]
    
    def _format_field_name(self, field: str) -> str:
        """Convert field_name to Field Name."""
        return field.replace("_", " ").title()


class InnovationContextService:
    """
    Service for managing innovation context data.
    Handles persistence and retrieval of innovation details.
    """
    
    def __init__(self):
        self._contexts: Dict[str, Dict[str, str]] = {}
        self._current_context: Optional[Dict[str, str]] = None
    
    def set_current(self, context: Dict[str, str]) -> None:
        """Set the current active context."""
        self._current_context = context
    
    def get_current(self) -> Optional[Dict[str, str]]:
        """Get the current active context."""
        return self._current_context
    
    def save_context(self, name: str, context: Dict[str, str]) -> None:
        """Save a named context."""
        self._contexts[name] = context
    
    def load_context(self, name: str) -> Optional[Dict[str, str]]:
        """Load a named context."""
        return self._contexts.get(name)
    
    def list_saved_contexts(self) -> list:
        """List all saved context names."""
        return list(self._contexts.keys())
    
    def clear_current(self) -> None:
        """Clear the current context."""
        self._current_context = None


# =============================================================================
# SERVICE INSTANCES
# =============================================================================

# Global service instances for application-wide access
prompt_service = PromptGeneratorService()
context_service = InnovationContextService()
