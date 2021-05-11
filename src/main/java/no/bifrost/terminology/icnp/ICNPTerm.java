package no.bifrost.terminology.icnp;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class ICNPTerm{
    private String code;
    private String axis;
    private String term;
    private String definition;
    private String snomedTerm;
}