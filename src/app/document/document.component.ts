// word-document.service.ts
import {Component} from '@angular/core';
import {DocumentService} from './word-documen.service';
import {DocumentService2} from './word-documen2.service';
import {DocumentEssaludService} from './word-documen-EsSalud.service';
import {DocumentEssalud2Service} from './word-documen-EsSalud-2.service';

@Component({
  selector: 'app-document',
  standalone: true,
  templateUrl: './document.component.html',
  styleUrl: './document.component.scss'
})
export class DocumentComponent {
  constructor(private documentService: DocumentService,
              private documentService2: DocumentService2,
              public documentEsSaludService: DocumentEssaludService,
              public documentEsSalud2Service: DocumentEssalud2Service) {
  }

  downloadDocument(): void {
    this.documentService.generateDocument();
  }
  downloadDocument2(): void {
    this.documentService2.generateDocument();
  }


  public documentEsSalud(): void {
    this.documentEsSaludService.generateDocument();
  }

  public documentEsSalud2(): void {
    this.documentEsSalud2Service.generateDocument2();
  }
}
